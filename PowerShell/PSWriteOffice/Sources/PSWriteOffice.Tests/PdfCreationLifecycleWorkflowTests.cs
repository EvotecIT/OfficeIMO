using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Cmdlets.Pdf;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PdfCreationLifecycleWorkflowTests {
    [Fact]
    public void LiveCreationDefersSavingAndBlocksPublishOnlyAfterSuccess() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-create-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "Review[final].pdf");
            using var runspace = Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("New-OfficePdf").AddParameter("Path", path).AddParameter("Theme", "Classic").AddParameter("Size", "11pt");
            using PdfDocument live = Assert.IsType<PdfDocument>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Equal(path, live.FilePath);
            Assert.False(File.Exists(path));
            shell.Commands.Clear();
            shell.AddCommand("Add-OfficePdfParagraph").AddParameter("Document", live).AddParameter("Text", "Live document evidence");
            shell.Invoke();
            Assert.False(shell.HadErrors, Errors(shell));
            live.Save(new PdfSaveOptions { AllowLoss = true });
            using (PdfDocument reopened = PdfDocument.Load(path)) Assert.Contains("Live document evidence", reopened.Read().Text);
            byte[] original = File.ReadAllBytes(path);

            runspace.SessionStateProxy.SetVariable("entered", false);
            shell.Commands.Clear();
            shell.AddCommand("New-OfficePdf").AddParameter("Path", path).AddParameter("AllowLoss")
                .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; PdfParagraph 'Rejected replacement'"));
            Exception? failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
            Assert.Equal(original, File.ReadAllBytes(path));

            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            shell.AddCommand("New-OfficePdf").AddParameter("Path", path).AddParameter("Force").AddParameter("AllowLoss")
                .AddParameter("Content", ScriptBlock.Create("PdfParagraph 'Discarded'; throw 'authoring-failed'"));
            failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.Contains("authoring-failed", failure?.Message ?? Errors(shell));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(root));

            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            shell.AddCommand("New-OfficePdf").AddParameter("Path", path).AddParameter("Force").AddParameter("WhatIf")
                .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; PdfParagraph 'Skipped'"));
            Assert.Empty(shell.Invoke());
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
            Assert.Equal(original, File.ReadAllBytes(path));

            shell.Commands.Clear();
            shell.AddCommand("New-OfficePdf").AddParameter("Path", path).AddParameter("Force").AddParameter("PassThru").AddParameter("AllowLoss")
                .AddParameter("Theme", "Modern").AddParameter("Size", "12pt")
                .AddParameter("Content", ScriptBlock.Create("PdfHeading 'Published'; PdfParagraph 'Completed content'"));
            Assert.Equal(path, Assert.IsType<FileInfo>(Assert.Single(shell.Invoke()).BaseObject).FullName);
            Assert.False(shell.HadErrors, Errors(shell));
            using (PdfDocument reopened = PdfDocument.Load(path)) Assert.Contains("Completed content", reopened.Read().Text);
            Assert.Single(Directory.GetFiles(root));
            byte[] published = File.ReadAllBytes(path);
            shell.Commands.Clear();
            shell.AddCommand("Close-OfficePdf").AddParameter("Document", live);
            Assert.Empty(shell.Invoke());
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Equal(published, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void TypedPermissionsOnlyCreationKeepsOwnerAuthorizationSeparateFromOpening() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-permissions-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "Restricted.pdf");
            using SecureString owner = new();
            foreach (char character in "owner-secret") owner.AppendChar(character);
            owner.MakeReadOnly();
            using var runspace = Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("param($path, $owner) New-OfficePdf -Path $path -OwnerPassword $owner -Permission Print,Accessibility -AllowLoss { PdfParagraph 'Owner-only extraction' }")
                .AddArgument(path).AddArgument(owner);
            Assert.Empty(shell.Invoke());
            Assert.False(shell.HadErrors, Errors(shell));
            using (PdfDocument reader = PdfDocument.Load(path)) Assert.Throws<PdfPermissionDeniedException>(() => reader.Read());
            using (PdfDocument authenticated = PdfDocument.Load(path, new PdfLoadOptions { Password = "owner-secret" }))
                Assert.Contains("Owner-only extraction", authenticated.Read().Text);
            shell.Commands.Clear();
            shell.AddCommand("New-OfficePdf").AddParameter("Permission", PdfStandardPermissions.Print);
            Exception? failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.Contains("requires Password or OwnerPassword", failure?.Message ?? Errors(shell));
        } finally { Directory.Delete(root, true); }
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(NewOfficePdfCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
}
