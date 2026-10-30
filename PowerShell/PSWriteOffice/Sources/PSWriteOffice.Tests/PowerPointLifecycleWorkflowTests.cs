using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PowerPointLifecycleWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-ppt-lifecycle-" + Guid.NewGuid().ToString("N"));
    public PowerPointLifecycleWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public void NewWithoutContentReturnsALiveDocumentWithoutWritingOnDispose() {
        string path = Path.Combine(_root, "created-later.pptx");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("New-OfficePowerPoint").AddParameter("Path", path);
        var result = Assert.Single(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using (var document = Assert.IsType<PowerPointPresentation>(result.BaseObject)) {
            Assert.Equal(path, document.FilePath);
            document.AddSlide().AddTextBoxPoints("Saved later", 20D, 20D, 200D, 40D);
            Assert.False(File.Exists(path));
        }
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void ForceBlockFailureRunsTheBlockAndPreservesTheExistingDestination() {
        string path = Path.Combine(_root, "existing.pptx");
        byte[] original = { 1, 2, 3, 4 };
        File.WriteAllBytes(path, original);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("entered", false);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("New-OfficePowerPoint").AddParameter("Path", path).AddParameter("Force")
            .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; throw 'authoring-failed'"));
        Exception? failure = Record.Exception(() => shell.Invoke());

        Assert.True((bool)runspace.SessionStateProxy.GetVariable("entered"));
        Assert.True(failure != null || shell.HadErrors);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Single(Directory.GetFiles(_root));
    }

    [Fact]
    public void WhatIfDoesNotRunTheContentBlockOrCreateAFile() {
        string path = Path.Combine(_root, "whatif.pptx");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("entered", false);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("New-OfficePowerPoint").AddParameter("Path", path).AddParameter("WhatIf")
            .AddParameter("Content", ScriptBlock.Create("$global:entered = $true"));
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
        Assert.Empty(Directory.GetFiles(_root));
    }

    [Fact]
    public void SuccessfulContentSavesAValidDeckAndReturnsOnlyTheFile() {
        string path = Path.Combine(_root, "saved.pptx");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("New-OfficePowerPoint").AddParameter("Path", path).AddParameter("PassThru")
            .AddParameter("Content", ScriptBlock.Create("Add-OfficePowerPointSlide -Title 'Created in a block'"));
        var result = Assert.Single(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(path, Assert.IsType<FileInfo>(result.BaseObject).FullName);
        using var saved = PowerPointPresentation.Load(path);
        Assert.Single(saved.Slides);
        Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void LiveNewSaveOpenAndCloseUseOneDocumentParameterAndLiteralInput() {
        string path = Path.Combine(_root, "Review[final].pptx");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        var document = Assert.IsType<PowerPointPresentation>(Assert.Single(shell.AddCommand("New-OfficePowerPoint").Invoke()).BaseObject);
        try {
            document.AddSlide().AddTextBoxPoints("Explicit lifecycle", 20D, 20D, 200D, 40D);
            shell.Commands.Clear();
            shell.AddCommand("Save-OfficePowerPoint").AddParameter("Document", document).AddParameter("Path", path).AddParameter("PassThru");
            Assert.Equal(path, Assert.IsType<FileInfo>(Assert.Single(shell.Invoke()).BaseObject).FullName);
            Assert.False(shell.HadErrors, Errors(shell));
            shell.Commands.Clear();
            Assert.Empty(shell.AddCommand("Close-OfficePowerPoint").AddParameter("Document", document).Invoke());
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Throws<ObjectDisposedException>(() => document.OpenXmlDocument);

            shell.Commands.Clear();
            using var opened = Assert.IsType<PowerPointPresentation>(Assert.Single(shell.AddCommand("Open-OfficePowerPoint")
                .AddParameter("LiteralPath", path).Invoke()).BaseObject);
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Contains(opened.Slides.SelectMany(slide => slide.TextBoxes), box => box.Text == "Explicit lifecycle");
        } finally { document.Dispose(); }
    }

    [Fact]
    public void OpeningAndSavingAnEncryptedDeckPreservesItsPasswordWithoutASaveOverride() {
        string path = Path.Combine(_root, "encrypted.pptx");
        string copy = Path.Combine(_root, "encrypted-copy.pptx");
        const string password = "synthetic-lifecycle-password";
        using (var source = PowerPointPresentation.Create()) {
            source.AddSlide().AddTextBoxPoints("Encrypted lifecycle", 20D, 20D, 200D, 40D);
            source.Save(path, new PowerPointSaveOptions { Password = password });
        }
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        using var document = Assert.IsType<PowerPointPresentation>(Assert.Single(shell.AddCommand("Open-OfficePowerPoint")
            .AddParameter("LiteralPath", path).AddParameter("Password", password).Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Single(shell.Streams.Warning);
        shell.Commands.Clear();
        shell.AddCommand("Save-OfficePowerPoint").AddParameter("Document", document).AddParameter("Path", copy).Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var reopened = PowerPointPresentation.Load(copy, new PowerPointLoadOptions { Password = password });
        Assert.Contains(reopened.Slides.SelectMany(slide => slide.TextBoxes), box => box.Text == "Encrypted lifecycle");
        Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, File.ReadAllBytes(copy).Take(4));
    }

    [Fact]
    public void FailedOpenBlockKeepsTheOriginalBytesAndCleansItsStagingFile() {
        string path = Path.Combine(_root, "existing-open.pptx");
        using (var source = PowerPointPresentation.Create()) source.Save(path);
        byte[] original = File.ReadAllBytes(path);
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("entered", false);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Open-OfficePowerPoint").AddParameter("LiteralPath", path)
            .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; throw 'authoring-failed'"));
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True((bool)runspace.SessionStateProxy.GetVariable("entered"));
        Assert.True(failure != null || shell.HadErrors);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Single(Directory.GetFiles(_root));
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficePowerPoint", typeof(NewOfficePowerPointCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Open-OfficePowerPoint", typeof(OpenOfficePowerPointCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Save-OfficePowerPoint", typeof(SaveOfficePowerPointCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Close-OfficePowerPoint", typeof(CloseOfficePowerPointCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficePowerPointSlide", typeof(AddOfficePowerPointSlideCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Set-OfficePowerPointSlide", typeof(SetOfficePowerPointSlideCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
}
