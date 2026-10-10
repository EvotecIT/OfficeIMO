using OfficeIMO.Pdf;
using OfficeIMO.Tool.Commands.Workflow;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class WorkflowCommandTests {
    [Theory]
    [InlineData(".odg")]
    [InlineData(".fodg")]
    public async Task DrawBatchUsesTheCanonicalPageConverter(string extension) {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "drawing" + extension), output = Path.Combine(scope.Path, "PDF");
        var drawing = OfficeIMO.OpenDocument.OdgDocument.Create();
        drawing.AddPage("Caption", OfficeIMO.OpenDocument.OdfLength.Points(360), OfficeIMO.OpenDocument.OdfLength.Points(220))
            .Shapes.AddTextBox(OfficeIMO.OpenDocument.OdfRect.FromCentimeters(1, 1, 5, 2), "ToolCaption");
        drawing.AddPage("Blank", OfficeIMO.OpenDocument.OdfLength.Points(180), OfficeIMO.OpenDocument.OdfLength.Points(120));
        if (extension == ".fodg") drawing.SaveFlatXml(input); else drawing.Save(input);
        ToolResult result = await RunAsync(["workflow", "batch", input, "--output", output, "--target", "pdf"]);
        Assert.Equal((int)OfficeImoToolExitCode.Success, result.ExitCode);
        string pdfPath = Assert.Single(Directory.GetFiles(output, "*.pdf"));
        var pdf = PdfReadDocument.Open(File.ReadAllBytes(pdfPath));
        Assert.Equal(2, pdf.Pages.Count);
        Assert.Equal((180D, 120D), pdf.Pages[1].GetPageSize());
        Assert.Contains("ToolCaption", pdf.ExtractText());
    }

    [Fact]
    public async Task WordImagesCommandPublishesSeparateWordAndPdfCopiesAndReportsAnalysis() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.docx"), output = Path.Combine(scope.Path, "output.docx"), pdf = Path.Combine(scope.Path, "output.pdf");
        using (var word = OfficeIMO.Word.WordDocument.Create()) {
            var image = new OfficeIMO.Drawing.OfficeRasterImage(200, 100, OfficeIMO.Drawing.OfficeColor.Red);
            using var stream = new MemoryStream(OfficeIMO.Drawing.OfficeJpegCodec.Encode(image));
            word.AddParagraph("Retained text").InsertImage(stream, "picture.jpg", 96, 48);
            word.Save(input);
        }
        byte[] before = File.ReadAllBytes(input);
        var analysis = await RunAsync(["workflow", "optimize-images", input, "--analyze"]);
        var copied = await RunAsync(["workflow", "optimize-images", input, "--output", output, "--dpi", "72"]);
        var exported = await RunAsync(["workflow", "optimize-images", input, "--output", pdf, "--mode", "recompress", "--quality", "40"]);
        Assert.Equal((int)OfficeImoToolExitCode.Success, analysis.ExitCode);
        Assert.Contains("WordImageInventory", analysis.Output);
        Assert.Equal((int)OfficeImoToolExitCode.Success, copied.ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.Success, exported.ExitCode);
        Assert.Equal(before, File.ReadAllBytes(input));
        Assert.Single(PdfReadDocument.Open(File.ReadAllBytes(pdf)).Pages);
        using var reopened = OfficeIMO.Word.WordDocument.Load(output);
        Assert.Equal(72, OfficeIMO.Drawing.OfficeImageReader.Identify(reopened.Images[0].ToBytes()).Width);
    }

    [Theory]
    [InlineData("--output", "--force")]
    [InlineData("--mode", "invalid")]
    [InlineData("--quality", "101")]
    public async Task WordImagesCommandRejectsInvalidOptionsBeforeFileAccess(string option, string value) {
        var result = await RunAsync(["workflow", "optimize-images", "source.docx", "--analyze", option, value]);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, result.ExitCode);
    }
    [Fact]
    public async Task ExportPagesPublishesSelectedValidatedImages() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.pdf");
        string output = Path.Combine(scope.Path, "pages");
        CreatePdf(input, 3);

        ToolResult result = await RunAsync([
            "workflow", "export-pages", input, "--output", output, "--pages", "3,1", "--format", "png"
        ]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, result.ExitCode);
        Assert.True(Directory.Exists(output));
        Assert.Equal(2, Directory.GetFiles(output, "*.png").Length);
        Assert.Contains("2 page image(s)", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain("failed", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AssembleUsesOrderedSourcesAndRefusesAnExistingDestinationWithoutForce() {
        using var scope = new TestDirectory();
        string first = Path.Combine(scope.Path, "first.pdf");
        string second = Path.Combine(scope.Path, "second.pdf");
        string output = Path.Combine(scope.Path, "assembled.pdf");
        CreatePdf(first, 1);
        CreatePdf(second, 2);

        ToolResult initial = await RunAsync(["workflow", "assemble", first, second, "--output", output]);
        Assert.Equal((int)OfficeImoToolExitCode.Success, initial.ExitCode);
        Assert.Equal(3, PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.Count);

        ToolResult refused = await RunAsync(["workflow", "assemble", second, "--output", output]);
        ToolResult replaced = await RunAsync(["workflow", "assemble", second, "--output", output, "--force"]);

        Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, refused.ExitCode);
        Assert.Contains("exists", refused.Error, StringComparison.OrdinalIgnoreCase);
        Assert.Equal((int)OfficeImoToolExitCode.Success, replaced.ExitCode);
        Assert.Equal(2, PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.Count);
    }

    [Fact]
    public async Task PrintPlanReportsSelectedPagesAndSheetComposition() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.pdf");
        CreatePdf(input, 4);

        ToolResult result = await RunAsync([
            "workflow", "print-plan", input, "--pages", "4,2,1", "--paper", "Letter",
            "--orientation", "landscape", "--pages-per-sheet", "2", "--scale", "actual"
        ]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, result.ExitCode);
        Assert.Contains("Selected pages: 4,2,1", result.Output, StringComparison.Ordinal);
        Assert.Contains("Sheets: 2", result.Output, StringComparison.Ordinal);
        Assert.Contains("Sheet 1: 792 x 612 pt; pages 4,2", result.Output, StringComparison.Ordinal);
        Assert.Equal(string.Empty, result.Error);
    }

    [Fact]
    public async Task ExtendedPrintPlanOptionsReachTheCanonicalPlanner() {
        var output = new StringWriter();
        var error = new StringWriter();
        PdfPrintPlanRequest? captured = null;
        int exit = await WorkflowCommand.RunAsync([
            "print-plan", "source.pdf", "--pages", "9,8,7,6,5,4,3,2,1", "--pages-per-sheet", "6",
            "--scale", "custom", "--custom-scale", "25", "--alignment", "bottom-right", "--page-subset", "odd",
            "--margin", "12", "--margin-left", "20", "--margin-top", "30", "--margin-right", "40", "--margin-bottom", "50", "--color", "grayscale"
        ], output, error, CancellationToken.None, printPlanner: (request, token) => {
            captured = request;
            PdfDocument document = PdfDocument.Create(c => { for (int i = 0; i < 9; i++) c.Page(p => p.Size(100, 120).Margin(0)); });
            return Task.FromResult(PdfPrintPlanner.Create(document, request, token));
        });
        Assert.Equal((int)OfficeImoToolExitCode.Success, exit);
        Assert.NotNull(captured);
        Assert.Equal(PdfPrintScaleMode.Custom, captured.ScaleMode);
        Assert.Equal(25, captured.CustomScalePercent);
        Assert.Equal(PdfPrintAlignment.BottomRight, captured.Alignment);
        Assert.Equal(20, captured.MarginLeft); Assert.Equal(30, captured.MarginTop);
        Assert.Equal(40, captured.MarginRight); Assert.Equal(50, captured.MarginBottom);
        Assert.Contains("Selected pages: 9,7,5,3,1", output.ToString());
        Assert.Contains("Color: Grayscale", output.ToString());
        Assert.Empty(error.ToString());
    }

    [Fact]
    public async Task InvalidWorkflowOptionsReturnSharedUsageCode() {
        ToolResult result = await RunAsync(["workflow", "assemble", "source.pdf", "--format", "png", "--output", "result.pdf"]);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, result.ExitCode);
        Assert.Contains("--format is not valid with assemble", result.Error, StringComparison.Ordinal);
        Assert.Contains("officeimo workflow assemble", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MissingWorkflowInputPreservesInputNotFoundExitCode() {
        using var scope = new TestDirectory();
        string missing = Path.Combine(scope.Path, "missing.pdf");
        string output = Path.Combine(scope.Path, "pages");

        ToolResult result = await RunAsync(["workflow", "export-pages", missing, "--output", output]);

        Assert.Equal((int)OfficeImoToolExitCode.InputNotFound, result.ExitCode);
        Assert.Contains("does not exist", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task PrintPlanAccessFailureIsNotClassifiedAsAnOutputFailure() {
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await WorkflowCommand.RunAsync(
            ["print-plan", "source.pdf"],
            output,
            error,
            printPlanner: static (_, _) => Task.FromException<PdfPrintPlan>(
                new UnauthorizedAccessException("Input denied.")));

        Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, exitCode);
        Assert.Contains("Input access failed", error.ToString(), StringComparison.Ordinal);

        error.GetStringBuilder().Clear();
        exitCode = await WorkflowCommand.RunAsync(
            ["print-plan", "source.pdf"],
            output,
            error,
            printPlanner: static (_, _) => Task.FromException<PdfPrintPlan>(
                new IOException("Input read failed.")));

        Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, exitCode);
        Assert.Contains("Input I/O failed", error.ToString(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("--output", "--force")]
    [InlineData("--pages", "--paper")]
    [InlineData("--format", "--force")]
    public async Task OptionTokenCannotBeConsumedAsAWorkflowOptionValue(string option, string nextOption) {
        ToolResult result = await RunAsync(["workflow", "export-pages", "source.pdf", option, nextOption]);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, result.ExitCode);
        Assert.Contains(option + " requires a value", result.Error, StringComparison.Ordinal);
    }

    private static async Task<ToolResult> RunAsync(string[] args) {
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        int exitCode = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);
        return new ToolResult(exitCode, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }

    private static void CreatePdf(string path, int pageCount) {
        PdfDocument.Create(document => {
            for (int page = 1; page <= pageCount; page++) {
                int pageNumber = page;
                document.Page(current => current
                    .Size(400 + pageNumber * 10, 600)
                    .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Page " + pageNumber)))));
            }
        }).Save(path);
    }

    private sealed record ToolResult(int ExitCode, string Output, string Error);

    private sealed class TestDirectory : IDisposable {
        internal TestDirectory() {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-tool-workflow-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        internal string Path { get; }

        public void Dispose() {
            try {
                Directory.Delete(Path, recursive: true);
            } catch (IOException) {
                // Best effort for transient package handles on Windows.
            } catch (UnauthorizedAccessException) {
                // Best effort for transient package handles on Windows.
            }
        }
    }
}
