using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using PSWriteOffice.Cmdlets.Pdf;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PdfOpenLifecycleWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-open-" + Guid.NewGuid().ToString("N"));
    public PdfOpenLifecycleWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public void LoadedBlockSavesSuccessiveEditsToAnOutputCopy() {
        string source = CreateSource("review[final].pdf"), output = Path.Combine(_root, "published.pdf");
        byte[] original = File.ReadAllBytes(source);
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("param($file, $output) $file | Open-OfficePdf -OutputPath $output -PassThru -Content { Remove-OfficePdfPage -Page 2; PdfMetadata -Title 'Reviewed'; Remove-OfficePdfPage -Page last }")
            .AddArgument(new FileInfo(source)).AddArgument(output);
        var results = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(output, Assert.IsType<FileInfo>(Assert.Single(results).BaseObject).FullName);
        Assert.Equal(original, File.ReadAllBytes(source));
        using var saved = PdfDocument.Load(output);
        Assert.Single(saved.Pages);
        Assert.Contains("First page", saved.Read().Text);
        Assert.DoesNotContain("Second page", saved.Read().Text);
        Assert.Equal("Reviewed", saved.Inspect().Metadata.Title);
        Assert.Equal(2, Directory.GetFiles(_root).Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RejectedOrFailedBlockPreservesBothFiles(bool whatIf) {
        string source = CreateSource("input.pdf"), output = CreateSource("existing.pdf");
        byte[] original = File.ReadAllBytes(source), destination = File.ReadAllBytes(output);
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        runspace.SessionStateProxy.SetVariable("entered", false);
        shell.AddCommand("Open-OfficePdf").AddParameter("LiteralPath", source).AddParameter("OutputPath", output)
            .AddParameter("Content", ScriptBlock.Create("$global:entered = $true; Remove-OfficePdfPage -Page 2; throw 'edit-failed'"));
        if (whatIf) shell.AddParameter("WhatIf");
        Exception? error = Record.Exception(() => shell.Invoke());
        Assert.Equal(!whatIf, (bool)runspace.SessionStateProxy.GetVariable("entered"));
        Assert.Equal(!whatIf, error != null || shell.HadErrors);
        Assert.Equal(original, File.ReadAllBytes(source));
        Assert.Equal(destination, File.ReadAllBytes(output));
        Assert.Equal(2, Directory.GetFiles(_root).Length);
    }

    [Fact]
    public void WildcardOpenReturnsLiveDocumentsAndCloseDoesNotWrite() {
        string first = CreateSource("first.pdf"), second = CreateSource("second.pdf");
        byte[] original = File.ReadAllBytes(first);
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Open-OfficePdf").AddParameter("Path", Path.Combine(_root, "*.pdf"));
        var documents = shell.Invoke().Select(value => Assert.IsType<PdfDocument>(value.BaseObject)).ToArray();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(2, documents.Length);
        foreach (PdfDocument document in documents) {
            Assert.Contains(document.FilePath, new[] { first, second });
            shell.Commands.Clear();
            shell.AddCommand("Close-OfficePdf").AddParameter("Document", document);
            Assert.Empty(shell.Invoke());
            Assert.False(shell.HadErrors, Errors(shell));
        }
        Assert.Equal(original, File.ReadAllBytes(first));
    }

    [Fact]
    public void MultiplePipelineInputsRefuseTheBlockBeforeReplacingOneOutput() {
        string first = CreateSource("first.pdf"), second = CreateSource("second.pdf"), output = CreateSource("existing.pdf");
        byte[] original = File.ReadAllBytes(output);
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        runspace.SessionStateProxy.SetVariable("entered", false);
        shell.AddScript("param($files, $output) $files | Open-OfficePdf -OutputPath $output -Content { $global:entered = $true; PdfMetadata -Title 'Must not publish' }")
            .AddArgument(new[] { new FileInfo(first), new FileInfo(second) }).AddArgument(output);
        Exception? error = Record.Exception(() => shell.Invoke());
        Assert.True(error != null || shell.HadErrors);
        Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
        Assert.Equal(original, File.ReadAllBytes(output));
        Assert.Equal(3, Directory.GetFiles(_root).Length);
    }

    [Fact]
    public void SecurePasswordOverrideRetainsTheNativePageSelection() {
        string path = CreateSource("restricted.pdf", "open-secret");
        var options = new PdfLoadOptions { Password = "do-not-use", PageSelection = PdfPageSelection.From(2) };
        using var password = new System.Security.SecureString();
        foreach (char character in "open-secret") password.AppendChar(character);
        password.MakeReadOnly();
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Open-OfficePdf").AddParameter("LiteralPath", path).AddParameter("LoadOptions", options)
            .AddParameter("Password", password);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        using var document = Assert.IsType<PdfDocument>(Assert.Single(result).BaseObject);
        var read = document.Read();
        Assert.Single(read.Pages);
        Assert.Contains("Second page", read.Text);
        Assert.DoesNotContain("First page", read.Text);
        Assert.DoesNotContain("Last page", read.Text);
        Assert.Equal("do-not-use", options.Password);
    }

    private string CreateSource(string name, string? password = null) {
        string path = Path.Combine(_root, name);
        using var document = PdfDocument.Create();
        document.Compose(compose => compose.Content(content => content.Paragraph(p => p.Text("First page")).PageBreak()
            .Paragraph(p => p.Text("Second page")).PageBreak().Paragraph(p => p.Text("Last page"))));
        document.Save(path, new PdfSaveOptions { AllowLoss = true, Encryption = password == null ? null : new PdfStandardEncryptionOptions(password) });
        return path;
    }
    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(NewOfficePdfCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error.Select(error => error.ToString()));
}
