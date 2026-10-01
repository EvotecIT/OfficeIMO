using System.Collections.Concurrent;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using ReadPdf = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeConversionBatchTests {
    [Fact]
    public async Task MixedFormatsUseTheCatalogAndReturnSkippedFilesAndVerifiedReuse() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "text.txt"), "Literal <h1> text");
        File.WriteAllText(Path.Combine(scope.Input, "markup.md"), "# Markdown source", System.Text.Encoding.Unicode);
        File.WriteAllText(Path.Combine(scope.Input, "page.html"), "<html><body><h1>HTML source</h1></body></html>");
        File.WriteAllText(Path.Combine(scope.Input, "rich.rtf"), "{\\rtf1\\ansi RTF source}");
        File.WriteAllText(Path.Combine(scope.Input, "other.bin"), "Unsupported input");
        using (var word = WordDocument.Create(Path.Combine(scope.Input, "word.docx"))) { word.AddParagraph("Word source"); word.Save(); }
        using (var excel = ExcelDocument.Create(Path.Combine(scope.Input, "sheet.xlsx"))) { excel.AddWorksheet("Data").Cell(1, 1, "Excel source"); excel.Save(); }
        using (var slides = PowerPointPresentation.Create(Path.Combine(scope.Input, "slides.pptx"))) { slides.AddSlide().AddTextBoxPoints("Slide source", 40, 40, 500, 60); slides.Save(); }
        var items = new Items();
        var builder = OfficeWorkflow.ConvertDirectory(scope.Input).ToDirectory(scope.Output).WithCheckpoint(scope.State)
            .WithConversionOptions(new OfficeWorkflowConversionOptions {
                Word = new OfficeIMO.Word.Pdf.WordToPdfOptions { Title = "Batch Word" },
                Excel = new OfficeIMO.Excel.Pdf.ExcelToPdfOptions { PdfOptions = new PdfOptions { DefaultFontSize = 12 } },
                PowerPoint = new OfficeIMO.PowerPoint.Pdf.PowerPointToPdfOptions { PdfOptions = new PdfOptions { DefaultFontSize = 12 } },
                Html = new OfficeIMO.Html.Pdf.HtmlToPdfOptions(),
                Markdown = new OfficeIMO.Markdown.Pdf.MarkdownToPdfOptions { Title = "Batch Markdown" },
                Rtf = new OfficeIMO.Rtf.Pdf.RtfToPdfOptions()
            });
        var first = await builder.RunAsync(progress: items);
        Assert.Equal(7, first.Completed); Assert.Equal(0, first.Failed); Assert.Equal(1, first.Skipped);
        Assert.Single(items.Values, item => item.Skipped);
        foreach (var item in items.Values.Where(item => !item.Skipped)) {
            Assert.Equal(OfficeWorkflowStatus.Completed, item.Status);
            using var pdf = ReadPdf.Open(item.OutputPath!);
            Assert.NotEmpty(pdf.GetPages());
        }
        var second = await builder.RunAsync();
        Assert.Equal(7, second.Reused); Assert.Equal(1, second.Skipped);
    }

    [Fact]
    public async Task EncryptedOfficeFilesUseTheExistingLoadersAndPasswordsAreNotCheckpointed() {
        using var scope = new BatchDirectory();
        const string password = "Batch qualification secret";
        string wordPath = Path.Combine(scope.Input, "locked.docx");
        string excelPath = Path.Combine(scope.Input, "locked.xlsx");
        using (var word = WordDocument.Create(wordPath)) { word.AddParagraph("Encrypted Word source"); word.SaveEncrypted(wordPath, password); }
        using (var excel = ExcelDocument.Create(excelPath)) { excel.AddWorksheet("Data").Cell(1, 1, "Encrypted Excel source"); excel.SaveEncrypted(excelPath, password); }
        var result = await OfficeWorkflow.ConvertDirectory(scope.Input).ToDirectory(scope.Output).WithCheckpoint(scope.State)
            .WithConversionOptions(new OfficeWorkflowConversionOptions { SourcePassword = password }).RunAsync();
        Assert.Equal(2, result.Completed); Assert.Equal(0, result.Failed);
        foreach (string checkpoint in Directory.EnumerateFiles(scope.State, "*.json", SearchOption.AllDirectories))
            Assert.DoesNotContain(password, File.ReadAllText(checkpoint));
    }

    [Fact]
    public async Task BatchTargetSelectionReusesTheExistingPdfToHtmlRoute() {
        using var scope = new BatchDirectory();
        string source = Path.Combine(scope.Input, "source.pdf");
        PdfPlainTextConverter.ToPdfDocumentResult("PDF batch source").Save(source);
        var result = await OfficeWorkflow.ConvertFiles(source).ToDirectory(scope.Output, ".html").WithCheckpoint(scope.State).RunAsync();
        Assert.Equal(1, result.Completed);
        Assert.Contains("PDF batch source", File.ReadAllText(Path.Combine(scope.Output, "source.pdf.html")));
    }

    [Fact]
    public async Task ExistingTextOptionsApplyToSingleAndBatchAndExecutionBudgetsDoNotInvalidateReuse() {
        using var scope = new BatchDirectory();
        string source = Path.Combine(scope.Input, "geometry.txt");
        File.WriteAllText(source, "literal geometry");
        var text = new PdfPlainTextOptions { PdfOptions = new PdfOptions {
            PageWidth = 240, PageHeight = 180, MarginLeft = 12, MarginRight = 12, MarginTop = 12, MarginBottom = 12,
            DefaultFont = PdfStandardFont.Courier, DefaultFontSize = 14
        } };
        var options = new OfficeWorkflowConversionOptions { PlainText = text };
        var single = await OfficeWorkflow.Convert(source).To(Path.Combine(scope.Output, "single.pdf")).WithConversionOptions(options).RunAsync();
        Assert.True(single.Succeeded);
        var request = new OfficeConversionBatchRequest { InputPaths = [source], OutputDirectory = scope.Output,
            CheckpointDirectory = scope.State, ConversionOptions = options, MaximumFiles = 1 };
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(request)).Completed);
        using (var pdf = ReadPdf.Open(Path.Combine(scope.Output, "geometry.txt.pdf"))) {
            Assert.Equal(240, pdf.GetPage(1).Width); Assert.Equal(180, pdf.GetPage(1).Height);
        }
        var resumed = await new OfficeWorkflowRunner().RunBatchAsync(request with {
            MaximumConcurrency = 9, MaximumFiles = 10, MaximumInputBytes = 128 * 1024 * 1024, MaximumOutputBytes = 512 * 1024 * 1024
        });
        Assert.Equal(1, resumed.Reused);
        text.PdfOptions.DefaultFontSize = 18;
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(request)).Failed);
    }

    [Fact]
    public async Task OrdinaryBatchesUseTheExistingRenameAndReplacePolicies() {
        using var scope = new BatchDirectory();
        string source = Path.Combine(scope.Input, "one.txt");
        File.WriteAllText(source, "First text");
        var builder = OfficeWorkflow.ConvertFiles(source).ToDirectory(scope.Output);
        Assert.Equal(1, (await builder.RunAsync()).Completed);
        Assert.Equal(1, (await builder.OnConflict(OfficeWorkflowConflictPolicy.Rename).RunAsync()).Completed);
        Assert.Equal(2, Directory.GetFiles(scope.Output, "*.pdf").Length);
        File.WriteAllText(source, "Replacement text");
        Assert.Equal(1, (await builder.OnConflict(OfficeWorkflowConflictPolicy.Replace).RunAsync()).Completed);
        using var pdf = ReadPdf.Open(Path.Combine(scope.Output, "one.txt.pdf"));
        Assert.Contains("Replacement text", pdf.GetPage(1).Text);
    }

    [Fact]
    public async Task HtmlCheckpointsIncludeLocalStylesheetContent() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "page.html"), "<html><head><link rel='stylesheet' href='style.css'></head><body><p>Styled source</p></body></html>");
        string css = Path.Combine(scope.Input, "style.css");
        File.WriteAllText(css, "p { font-size: 18px; }");
        var builder = OfficeWorkflow.ConvertDirectory(scope.Input).ToDirectory(scope.Output).WithCheckpoint(scope.State);
        Assert.Equal(1, (await builder.RunAsync()).Completed);
        Assert.Equal(1, (await builder.RunAsync()).Reused);
        byte[] original = File.ReadAllBytes(Path.Combine(scope.Output, "page.html.pdf"));
        File.WriteAllText(css, "p { font-size: 28px; }");
        Assert.Equal(1, (await builder.RunAsync()).Failed);
        Assert.Equal(original, File.ReadAllBytes(Path.Combine(scope.Output, "page.html.pdf")));
    }

    [Fact]
    public async Task DirectoryFilteringIsExplicitAndLiteralTextRemainsTheDefault() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "<h1>literal text</h1>");
        File.WriteAllText(Path.Combine(scope.Input, "two.md"), "# Unselected");
        Directory.CreateDirectory(Path.Combine(scope.Input, "nested"));
        File.WriteAllText(Path.Combine(scope.Input, "nested", "three.txt"), "Nested text");
        var result = await OfficeWorkflow.ConvertDirectory(scope.Input).ToDirectory(scope.Output).SelectExtensions(false, ".txt").RunAsync();
        Assert.Equal(1, result.Completed); Assert.Equal(1, result.Skipped);
        using var pdf = ReadPdf.Open(Path.Combine(scope.Output, "one.txt.pdf"));
        Assert.Contains("<h1>literal text</h1>", pdf.GetPage(1).Text);
    }

    private sealed class Items : IProgress<OfficeConversionBatchItemResult> {
        public ConcurrentBag<OfficeConversionBatchItemResult> Values { get; } = [];
        public void Report(OfficeConversionBatchItemResult item) => Values.Add(item);
    }
    private sealed class BatchDirectory : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-conversion-batch-" + Guid.NewGuid().ToString("N"));
        public string Input => Path.Combine(_root, "input");
        public string Output => Path.Combine(_root, "output");
        public string State => Path.Combine(_root, "state");
        public BatchDirectory() => Directory.CreateDirectory(Input);
        public void Dispose() => Directory.Delete(_root, recursive: true);
    }
}
