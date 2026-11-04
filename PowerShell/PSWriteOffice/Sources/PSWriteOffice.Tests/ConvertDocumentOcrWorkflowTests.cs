using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text.Json;
using OfficeIMO;
using OfficeIMO.CSV;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Cmdlets.Conversion;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services.Ocr;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ConvertDocumentOcrWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-convert-ocr-" + Guid.NewGuid().ToString("N"));
    public ConvertDocumentOcrWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, recursive: true);

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".docx")]
    [InlineData(".html")]
    [InlineData(".pdf")]
    [InlineData(".txt")]
    [InlineData(".md")]
    [InlineData(".json")]
    [InlineData(".csv")]
    public async Task ReviewPipelinePublishesAllLegacyProjectionsWithoutRereadingItsDeletedSource(string extension) {
        string source = Image("original.png"), output = Path.Combine(_root, "published", "ledger" + extension);
        int calls = 0;
        var engine = new DelegateOcrEngine("canonical-table", (_, _) => { calls++; return Task.FromResult(Table()); });
        var review = await PdfImageDocumentSource.FromFile(source).PrepareSearchableOcrAsync(engine, new PdfOcrMergeOptions { ReconstructLayout = true });
        File.Delete(source);
        var settings = new OfficeConversionOptions { PdfExcel = new() { UseFirstRowAsHeader = true } };
        using var runspace = Open();
        using var shell = Shell(runspace);
        shell.AddCommand("Convert-OfficeDocument").AddParameter("OutputPath", output).AddParameter("Options", settings).AddParameter("WhatIf");
        Assert.Empty(shell.Invoke(new object[] { review }));
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.False(Directory.Exists(Path.GetDirectoryName(output)));
        shell.Commands.Clear();
        shell.AddCommand("Convert-OfficeDocument").AddParameter("OutputPath", output).AddParameter("Options", settings);
        var result = Assert.IsType<OfficeImageConversionResult>(Assert.Single(shell.Invoke(new object[] { review })).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Same(review, result.Review);
        Assert.Equal(output, result.OutputPath);
        Assert.True(result.File.Exists);
        Assert.Equal(1, calls);
        Assert.True(settings.PdfExcel.UseFirstRowAsHeader);
        Assert.Contains(result.Review.Ocr.Pages.SelectMany(page => page.ProviderDiagnostics), diagnostic => diagnostic.Code == "provider-evidence");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "provider-evidence" && diagnostic.Severity == OfficeDiagnosticSeverity.Information);
        switch (extension) {
            case ".xlsx":
                using (var excel = ExcelDocument.Load(output)) Assert.Equal(-13.5, Assert.Single(excel.Sheets).Cell(3, 2).GetValue<double>());
                Assert.NotNull(result.TableScope);
                break;
            case ".docx":
                using (var zip = System.IO.Compression.ZipFile.OpenRead(output))
                using (var reader = new StreamReader(zip.GetEntry("word/document.xml")!.Open())) Assert.Contains("Service", reader.ReadToEnd());
                break;
            case ".pdf": using (var pdf = PdfDocument.Load(output)) Assert.Contains("Service", pdf.Read().Text); break;
            case ".json": using (var json = JsonDocument.Parse(File.ReadAllText(output))) Assert.Contains("Service", json.RootElement.GetRawText()); break;
            case ".csv": Assert.Contains("Credit,'-13.50", File.ReadAllText(output)); Assert.NotNull(result.TableScope); break;
            default: Assert.Contains("Service", File.ReadAllText(output)); break;
        }
        Assert.Single(Directory.GetFiles(Path.GetDirectoryName(output)!));
    }

    [Fact]
    public void OcrOptionsImpliesRecognitionAndWhatIfNeverCallsProviderOrPublishes() {
        string source = Image("scan.jp2"), output = Path.Combine(_root, "published", "review.csv");
        int calls = 0;
        var recognition = new OfficeOcrOptions { Provider = new DelegateOcrEngine("canonical-source", (_, _) => { calls++; return Task.FromResult(Table()); }) };
        var csv = new CsvSaveOptions { Delimiter = ';', FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Preserve };
        using var runspace = Open();
        using var shell = Shell(runspace);
        shell.AddCommand("Convert-OfficeDocument").AddParameter("Path", source).AddParameter("OutputPath", output)
            .AddParameter("OcrOptions", recognition).AddParameter("Options", csv).AddParameter("Show").AddParameter("WhatIf");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(0, calls);
        Assert.False(Directory.Exists(Path.GetDirectoryName(output)));
        shell.Commands.Clear();
        shell.AddCommand("Convert-OfficeDocument").AddParameter("Path", source).AddParameter("OutputPath", output)
            .AddParameter("OcrOptions", recognition).AddParameter("Options", csv);
        var result = Assert.IsType<OfficeImageConversionResult>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(1, calls);
        Assert.Contains("Credit;-13.50", File.ReadAllText(output));
        Assert.True(csv.IncludeHeader);
        Assert.Equal(CsvFormulaInjectionPolicy.Preserve, csv.FormulaInjectionPolicy);
        Assert.True(result.Review.Ocr.HasAcceptedOcrContent);
    }

    [Fact]
    public void PdfPageAndDpiOverrideCapturedRecognitionSettingsWithoutMutatingCaller() {
        string source = Path.Combine(_root, "two-pages.pdf"), output = Path.Combine(_root, "page.txt");
        byte[] image = File.ReadAllBytes(Image("scan.png"));
        File.WriteAllBytes(source, PdfDocument.Create(pdf => pdf.Content(content => content.AddImage(image, OfficeLength.Points(400), OfficeLength.Points(200)).PageBreak().AddImage(image, OfficeLength.Points(400), OfficeLength.Points(200)))).ToBytes());
        var requests = new List<OcrRequest>();
        var recognition = new OfficeOcrOptions { Provider = new DelegateOcrEngine("canonical-pages", (request, _) => { requests.Add(request); return Task.FromResult(Table()); }) };
        recognition.RecognitionOptions.ReadOptions.PageSelection = PdfPageSelection.From(1);
        recognition.RecognitionOptions.Dpi = 150;
        using var runspace = Open();
        using var shell = Shell(runspace);
        shell.AddCommand("Convert-OfficeDocument").AddParameter("Path", source).AddParameter("OutputPath", output)
            .AddParameter("OcrOptions", recognition).AddParameter("Page", "2").AddParameter("Dpi", 96D);
        var result = Assert.IsType<OfficeImageConversionResult>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        var page = Assert.Single(result.Review.Ocr.Pages);
        Assert.Equal(2, page.PageNumber);
        var request = Assert.Single(requests);
        Assert.Equal(2, request.PageNumber);
        Assert.Equal((int)Math.Ceiling(result.Review.GetPageSize(2).Width * 96D / 72D), request.PixelWidth);
        Assert.Equal("1", recognition.RecognitionOptions.ReadOptions.PageSelection.ToString());
        Assert.Equal(150D, recognition.RecognitionOptions.Dpi);
        Assert.Contains("Service", File.ReadAllText(output));
    }

    [Fact]
    public void NativeAdmissionAndOutputFailuresPreserveDestinationWithoutProvisioningOnRejectedSources() {
        string source = Image("scan.png"), output = Path.Combine(_root, "existing.txt");
        File.WriteAllText(output, "valuable destination");
        DateTime timestamp = new(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        File.SetLastWriteTimeUtc(output, timestamp);
        int calls = 0;
        var recognition = new OfficeOcrOptions { Provider = new DelegateOcrEngine("canonical-safety", (_, _) => { calls++; return Task.FromResult(Table()); }) };
        using var runspace = Open();
        using var shell = Shell(runspace);
        foreach (string mode in new[] { "InputLimit", "InvalidSource", "NoClobber", "OutputLimit" }) {
            shell.Commands.Clear(); shell.Streams.Error.Clear();
            string input = source;
            if (mode == "InvalidSource") { input = Path.Combine(_root, "not-an-image.bin"); File.WriteAllText(input, "unsupported"); }
            shell.AddCommand("Convert-OfficeDocument").AddParameter("Path", input).AddParameter("OutputPath", output).AddParameter("OcrOptions", recognition);
            if (mode == "InputLimit") shell.AddParameter("MaximumInputBytes", 1L);
            if (mode == "OutputLimit") shell.AddParameter("MaximumOutputBytes", 1L);
            if (mode == "NoClobber") shell.AddParameter("NoClobber");
            Assert.Throws<CmdletInvocationException>(() => shell.Invoke());
            Assert.Equal("valuable destination", File.ReadAllText(output));
            Assert.Equal(timestamp, File.GetLastWriteTimeUtc(output));
            Assert.Equal(mode is "NoClobber" ? 1 : mode is "OutputLimit" ? 2 : 0, calls);
        }
        Assert.DoesNotContain(Directory.GetFiles(_root), path => Path.GetFileName(path).Contains(".tmp"));
    }

    [Fact]
    public void OcrBatchUsesNativeImageExtensionsAndRetainsNoClobberSkipCounts() {
        string input = Path.Combine(_root, "inputs"), output = Path.Combine(_root, "outputs");
        Directory.CreateDirectory(input); Directory.CreateDirectory(output);
        File.Copy(Image("image.png"), Path.Combine(input, "first.j2k"));
        File.Copy(Path.Combine(_root, "image.png"), Path.Combine(input, "second.j2c"));
        string preserved = Path.Combine(output, "second.txt"); File.WriteAllText(preserved, "keep me");
        int calls = 0;
        var recognition = new OfficeOcrOptions { Provider = new DelegateOcrEngine("canonical-batch", (_, _) => { Interlocked.Increment(ref calls); return Task.FromResult(Table()); }) };
        using var runspace = Open();
        using var shell = Shell(runspace);
        shell.AddCommand("Convert-OfficeDocument").AddParameter("Path", input).AddParameter("OutputDirectory", output).AddParameter("To", "Text")
            .AddParameter("OcrOptions", recognition).AddParameter("NoClobber").AddParameter("SummaryVariable", "summary");
        var results = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(2, results.Count);
        var summary = Assert.IsType<OfficeBatchSummary>(runspace.SessionStateProxy.GetVariable("summary"));
        Assert.Equal(2, summary.Selected); Assert.Equal(1, summary.Completed); Assert.Equal(1, summary.Skipped); Assert.Equal(0, summary.Failed);
        Assert.Equal(1, calls);
        Assert.Contains("Service", File.ReadAllText(Path.Combine(output, "first.txt")));
        Assert.Equal("keep me", File.ReadAllText(preserved));
    }

    [Fact]
    public async Task ReviewAsStringKeepsNativeCsvPolicyAndOcrCannotLeakProviderIntoSavedJob() {
        var review = await new PdfImageDocumentSource(File.ReadAllBytes(Image("scan.png"))).PrepareSearchableOcrAsync(
            new DelegateOcrEngine("canonical-text", (_, _) => Task.FromResult(Table())), new PdfOcrMergeOptions { ReconstructLayout = true });
        var csv = new CsvSaveOptions { Delimiter = ';' };
        using var runspace = Open();
        using var shell = Shell(runspace);
        shell.AddCommand("Convert-OfficeDocument").AddParameter("Document", review).AddParameter("To", "Csv").AddParameter("AsString").AddParameter("Options", csv);
        string text = Assert.IsType<string>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Contains("Credit;-13.50", text); Assert.True(csv.IncludeHeader);
        shell.Commands.Clear();
        string job = Path.Combine(_root, "recognition.officejob.json");
        shell.AddCommand("Convert-OfficeDocument").AddParameter("Path", Path.Combine(_root, "scan.png")).AddParameter("To", "Text")
            .AddParameter("OutputDirectory", Path.Combine(_root, "outputs")).AddParameter("Ocr").AddParameter("SaveJob", job);
        var error = Assert.Throws<CmdletInvocationException>(() => shell.Invoke());
        Assert.IsType<PSArgumentException>(error.InnerException);
        Assert.False(File.Exists(job)); Assert.False(Directory.Exists(Path.Combine(_root, "outputs")));
    }

    private string Image(string name) {
        string path = Path.Combine(_root, name);
        File.WriteAllBytes(path, OfficeRasterImageEncoder.Encode(new OfficeRasterImage(800, 400, OfficeColor.White), OfficeImageExportFormat.Png));
        return path;
    }
    private static OcrResult Table() => new() {
        Diagnostics = new[] { new OcrDiagnostic { Code = "provider-evidence", Message = "Recognized with synthetic provider.", Severity = OcrDiagnosticSeverity.Info } },
        Spans = new[] { Word("Description", 20, 20, 75), Word("Amount", 200, 20, 45), Word("Service", 20, 40, 45), Word("119.88", 200, 40, 40),
            Word("Credit", 20, 60, 40), Word("-13.50", 200, 60, 40), Word("Plan", 20, 80, 30), Word("89.00", 200, 80, 40) }
    };
    private static OcrTextSpan Word(string text, double x, double y, double width) => new() {
        Text = text, Confidence = 1, Level = OcrTextSpanLevel.Word, CoordinateUnit = OcrCoordinateUnit.Points,
        Region = new OcrRegion { X = x, Y = y, Width = width, Height = 10 }
    };
    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open(); return runspace;
    }
    private static PowerShell Shell(Runspace runspace) { var shell = PowerShell.Create(); shell.Runspace = runspace; return shell; }
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
}
