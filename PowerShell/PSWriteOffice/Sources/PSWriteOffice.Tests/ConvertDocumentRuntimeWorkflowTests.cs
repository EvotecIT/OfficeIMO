using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ConvertDocumentRuntimeWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-convert-runtime-" + Guid.NewGuid().ToString("N"));
    public ConvertDocumentRuntimeWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, recursive: true);

    [Theory]
    [InlineData("Path")]
    [InlineData("LiteralPath")]
    [InlineData("StringPipeline")]
    [InlineData("FileInfoPipeline")]
    [InlineData("Text")]
    [InlineData("DocumentPipeline")]
    public void NoClobberPreservesDestinationsForFileAndMemoryInputs(string inputShape) {
        string source = Path.Combine(_root, "source.txt"), output = Path.Combine(_root, "output.pdf");
        File.WriteAllText(source, "Replacement evidence");
        byte[] original = { 1, 2, 3, 4 };
        File.WriteAllBytes(output, original);
        DateTime timestamp = new(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        File.SetLastWriteTimeUtc(source, timestamp.AddHours(-1));
        File.SetLastWriteTimeUtc(output, timestamp);
        using var document = MarkdownDocument.Parse("# Unsaved source");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("OutputPath", output).AddParameter("NoClobber");
        object[]? pipeline = null;
        switch (inputShape) {
            case "Path": case "LiteralPath":
                shell.AddParameter(inputShape, source).AddParameter("SkipUpToDate");
                break;
            case "StringPipeline":
                shell.AddParameter("From", "Text").AddParameter("SkipUpToDate");
                pipeline = new object[] { source };
                break;
            case "FileInfoPipeline":
                shell.AddParameter("From", "Text").AddParameter("SkipUpToDate");
                pipeline = new object[] { new FileInfo(source) };
                break;
            case "Text":
                shell.AddParameter("Text", "Literal replacement").AddParameter("From", "Text");
                break;
            case "DocumentPipeline":
                shell.AddParameter("From", "Markdown");
                pipeline = new object[] { document };
                break;
        }
        if (inputShape == "Text" || inputShape == "DocumentPipeline") {
            var failure = Assert.Throws<CmdletInvocationException>(() => shell.Invoke(pipeline));
            var conversion = Assert.IsType<OfficeConversionException>(failure.InnerException);
            Assert.IsType<IOException>(conversion.InnerException);
            Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ConversionOutputFailed);
        } else {
            var skipped = Assert.IsType<OfficeBatchItemReport>(Assert.Single(shell.Invoke(pipeline)).BaseObject);
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.True(skipped.Skipped);
            Assert.Equal("OutputExists", skipped.SkippedReason);
            Assert.Null(skipped.OutputPath);
            Assert.Null(skipped.Report);
        }
        Assert.Equal(original, File.ReadAllBytes(output));
        Assert.Equal(timestamp, File.GetLastWriteTimeUtc(output));
        Assert.Equal("Replacement evidence", File.ReadAllText(source));
        Assert.Single(document.Blocks);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void LiteralTextCreatesOneCompletePdfAndWhatIfPublishesNothing(bool pipeline) {
        string output = Path.Combine(_root, "output", "report.pdf");
        string file = Path.Combine(_root, "existing.html");
        File.WriteAllText(file, "File content must not replace literal input");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("output", output);
        runspace.SessionStateProxy.SetVariable("root", _root);
        runspace.SessionStateProxy.SetVariable("records", new[] { "<h1>Pipeline HTML</h1>", "<p>Full document body.</p>", "existing.html" });
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        string invoke = pipeline
            ? "$records | Convert-OfficeDocument -Text { $_ } -From Html -To Pdf -OutputPath $output"
            : "Convert-OfficeDocument -Text '<h1>Pipeline HTML</h1><p>Full document body.</p>existing.html' -From Html -To Pdf -OutputPath $output";
        shell.AddScript("Set-Location -LiteralPath $root; " + invoke + " -WhatIf -ErrorAction Stop");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.False(Directory.Exists(Path.GetDirectoryName(output)));
        shell.Commands.Clear();
        shell.AddScript(invoke + " -ErrorAction Stop");
        var report = Assert.IsType<OfficeConversionReport>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(output, report.OutputPath);
        string text = PdfDocument.Load(output).Read().Text;
        Assert.Contains("Pipeline HTML", text);
        Assert.Contains("Full document body.", text);
        Assert.Contains("existing.html", text);
        Assert.DoesNotContain("File content must not replace literal input", text);
        Assert.Equal("File content must not replace literal input", File.ReadAllText(file));
        Assert.Single(Directory.GetFiles(Path.GetDirectoryName(output)!));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FilePipelinesKeepFromOverridesInsteadOfTreatingPathsAsContent(bool fileInfo) {
        string source = Path.Combine(_root, "source.txt"), output = Path.Combine(_root, "output.docx");
        const string html = "<h1>File heading</h1><p>From override body.</p>";
        File.WriteAllText(source, html);
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument").AddParameter("From", "Html").AddParameter("To", "Word").AddParameter("OutputPath", output);
        object input = fileInfo ? new FileInfo(source) : source;
        Assert.IsType<OfficeConversionReport>(Assert.Single(shell.Invoke(new[] { input })).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        using var word = WordDocument.Load(output);
        Assert.Contains(word.Paragraphs, paragraph => paragraph.Text == "File heading");
        Assert.Contains(word.Paragraphs, paragraph => paragraph.Text == "From override body.");
        Assert.Equal(html, File.ReadAllText(source));
    }

    [Fact]
    public void NativeDocumentPipelineCapturesUnsavedContentWithoutClosingItsSource() {
        string output = Path.Combine(_root, "live.pdf");
        using var document = MarkdownDocument.Parse("# Live heading\n\nUnsaved body.");
        int blocks = document.Blocks.Count;
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument").AddParameter("From", "Markdown").AddParameter("To", "Pdf").AddParameter("OutputPath", output);
        Assert.IsType<OfficeConversionReport>(Assert.Single(shell.Invoke(new object[] { document })).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        string text = PdfDocument.Load(output).Read().Text;
        Assert.Contains("Live heading", text);
        Assert.Contains("Unsaved body.", text);
        Assert.Equal(blocks, document.Blocks.Count);
        document.P("Still live");
        Assert.Equal(blocks + 1, document.Blocks.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AccumulatedTextLimitRejectsBeforeAnyOutputIsPublished(bool commonOverride) {
        string output = Path.Combine(_root, "output", "limited.pdf");
        using var runspace = Open();
        runspace.SessionStateProxy.SetVariable("output", output);
        runspace.SessionStateProxy.SetVariable("settings", new OfficeConversionOptions { MaximumInputBytes = commonOverride ? 64 : 32 });
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        string limit = commonOverride ? " -MaximumInputBytes 32" : string.Empty;
        shell.AddScript("@('1234567890123456', '1234567890123456') | Convert-OfficeDocument -Text { $_ } -From Text -To Pdf -OutputPath $output -Options $settings" + limit + " -ErrorAction Stop");
        Assert.Empty(shell.Invoke());
        var failure = Assert.Single(shell.Streams.Error);
        Assert.IsType<PSArgumentException>(failure.Exception);
        Assert.Contains("MaximumInputBytes", failure.Exception.Message);
        Assert.False(Directory.Exists(Path.GetDirectoryName(output)));
    }

    [Fact]
    public void ShowIsAnOptionalSwitchOnTheSixDocumentConversionSets() {
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Get-Command").AddParameter("Name", "Convert-OfficeDocument");
        var command = Assert.IsType<CmdletInfo>(Assert.Single(shell.Invoke()).BaseObject);
        var show = command.Parameters[nameof(ConvertOfficeDocumentCommand.Show)];
        Assert.Equal(typeof(SwitchParameter), show.ParameterType);
        string[] sets = { "Path", "PathFile", "LiteralPath", "LiteralPathFile", "Text", "Document" };
        Assert.Equal(sets.OrderBy(value => value), show.ParameterSets.Keys.OrderBy(value => value));
        Assert.All(show.ParameterSets.Values, set => Assert.False(set.IsMandatory));
    }

    [Theory]
    [InlineData("File")]
    [InlineData("Text")]
    [InlineData("Document")]
    [InlineData("Batch")]
    public void ShowWithWhatIfDoesNotPublishAnyOutput(string inputShape) {
        string source = Path.Combine(_root, "source.txt");
        File.WriteAllText(source, "Show WhatIf source");
        string directory = Path.Combine(_root, "output"), output = Path.Combine(directory, "result.pdf");
        using var document = MarkdownDocument.Parse("# Show WhatIf source");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Show").AddParameter("WhatIf");
        switch (inputShape) {
            case "File": shell.AddParameter("Path", source).AddParameter("OutputPath", output); break;
            case "Text": shell.AddParameter("Text", "Show WhatIf source").AddParameter("From", "Text").AddParameter("OutputPath", output); break;
            case "Document": shell.AddParameter("Document", document).AddParameter("OutputPath", output); break;
            case "Batch": shell.AddParameter("Path", source).AddParameter("OutputDirectory", directory); break;
        }
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.False(Directory.Exists(directory));
        Assert.Equal("Show WhatIf source", File.ReadAllText(source));
    }

    [Theory]
    [InlineData("AsString")]
    [InlineData("NoOutputText")]
    [InlineData("NoOutputDocument")]
    [InlineData("Plan")]
    [InlineData("SaveJob")]
    public void ShowRejectsModesWithoutPublishedConversionFilesEvenWithWhatIf(string mode) {
        string source = Path.Combine(_root, "source.md");
        File.WriteAllText(source, "# Show source");
        string directory = Path.Combine(_root, "output"), job = Path.Combine(_root, "conversion.officejob.json");
        using var document = MarkdownDocument.Parse("# Show source");
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument").AddParameter("Show").AddParameter("WhatIf");
        switch (mode) {
            case "AsString": shell.AddParameter("Path", source).AddParameter("To", "Html").AddParameter("AsString"); break;
            case "NoOutputText": shell.AddParameter("Text", "# Show source").AddParameter("From", "Markdown").AddParameter("To", "Html"); break;
            case "NoOutputDocument": shell.AddParameter("Document", document).AddParameter("To", "Html"); break;
            default:
                shell.AddParameter("Path", source).AddParameter("To", "Pdf").AddParameter("OutputDirectory", directory);
                if (mode == "Plan") shell.AddParameter("Plan");
                else shell.AddParameter("SaveJob", job);
                break;
        }
        var failure = Assert.Throws<CmdletInvocationException>(() => shell.Invoke());
        Assert.IsType<PSArgumentException>(failure.InnerException);
        Assert.Contains("Show requires file output", failure.Message);
        Assert.False(Directory.Exists(directory));
        Assert.False(File.Exists(job));
        Assert.Equal("# Show source", File.ReadAllText(source));
    }

    [Theory]
    [InlineData("Path")]
    [InlineData("Text")]
    [InlineData("Document")]
    public void NativeOutputFailureRetainsWriteErrorCategoryAndRequestedDestination(string inputShape) {
        string source = Path.Combine(_root, "source.pdf");
        string blocker = Path.Combine(_root, "not-a-directory");
        string output = Path.Combine(blocker, "result.docx");
        File.WriteAllText(blocker, "valuable parent file");
        using (var pdf = PdfDocument.Create(pdf => pdf.Content(content => content.AddParagraph("Shared output failure")))) pdf.Save(source);
        using var live = PdfDocument.Load(source);
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Word").AddParameter("OutputPath", output);
        if (inputShape == "Path") shell.AddParameter("Path", source);
        else if (inputShape == "Text") shell.AddParameter("Text", "<p>Shared output failure</p>").AddParameter("From", "Html");
        else shell.AddParameter("Document", live);
        var failure = Assert.Throws<CmdletInvocationException>(() => shell.Invoke());
        Assert.Equal(ErrorCategory.WriteError, failure.ErrorRecord.CategoryInfo.Category);
        Assert.Equal(output, failure.ErrorRecord.TargetObject);
        var native = Assert.IsType<OfficeConversionException>(failure.InnerException);
        Assert.IsAssignableFrom<IOException>(native.InnerException);
        Assert.Contains(native.Report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ConversionOutputFailed);
        Assert.Equal("valuable parent file", File.ReadAllText(blocker));
        Assert.False(File.Exists(output));
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
}
