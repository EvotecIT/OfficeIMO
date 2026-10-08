using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class DrawPdfCommandTests {
    [Theory]
    [InlineData(".odg")]
    [InlineData(".fodg")]
    public async Task ConvertDrawPublishesReadablePagesAndLocatedLossEvidence(string extension) {
        string directory = CreateDirectory();
        try {
            string input = Path.Combine(directory, "source" + extension), destination = Path.Combine(directory, "result.pdf");
            CreateDrawing(input);
            byte[] original = File.ReadAllBytes(input);
            await using var output = new MemoryStream();
            using var error = new StringWriter();

            int exit = await OfficeImoToolApp.RunAsync(["convert", input, destination], Stream.Null, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exit);
            var pdf = PdfReadDocument.Open(File.ReadAllBytes(destination));
            Assert.Equal(2, pdf.Pages.Count);
            Assert.Equal((360D, 220D), pdf.Pages[0].GetPageSize());
            Assert.Equal((180D, 120D), pdf.Pages[1].GetPageSize());
            Assert.Contains("Draw CLI caption", pdf.ExtractText(), StringComparison.Ordinal);
            Assert.Equal(destination + Environment.NewLine, Encoding.UTF8.GetString(output.ToArray()));
            Assert.Contains("ODF_SKIPPED", error.ToString(), StringComparison.Ordinal);
            Assert.Contains("[document-metadata]", error.ToString(), StringComparison.Ordinal);
            Assert.Equal(original, File.ReadAllBytes(input));
            Assert.Empty(Directory.EnumerateFiles(directory, "*.tmp"));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(null, "PrintOnly", "ScreenOnly")]
    [InlineData("print", "PrintOnly", "ScreenOnly")]
    [InlineData("screen", "ScreenOnly", "PrintOnly")]
    public async Task DiagramLayerIntentSelectsRenderedContent(string? layers, string included, string excluded) {
        string directory = CreateDirectory();
        try {
            string input = Path.Combine(directory, "source.fodg"), destination = Path.Combine(directory, "result.pdf");
            var drawing = OdgDocument.Create();
            var page = drawing.AddPage("Layers", OdfLength.Points(360), OdfLength.Points(220));
            page.Layers.Add("Screen", OdgLayerDisplay.Screen);
            page.Layers.Add("Print", OdgLayerDisplay.Printer);
            page.Shapes.AddTextBox(Bounds(20), "ScreenOnly").Layer = "Screen";
            page.Shapes.AddTextBox(Bounds(100), "PrintOnly").Layer = "Print";
            drawing.SaveFlatXml(input);
            string[] args = layers is null ? ["convert", input, destination]
                : ["convert", input, destination, "--diagram-layers", layers];
            await using var output = new MemoryStream();
            using var error = new StringWriter();

            int exit = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.Success, exit);
            string text = PdfReadDocument.Open(File.ReadAllBytes(destination)).ExtractText();
            Assert.Contains(included, text, StringComparison.Ordinal);
            Assert.DoesNotContain(excluded, text, StringComparison.Ordinal);
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(".odg")]
    [InlineData(".fodg")]
    public async Task StrictLossRejectionRetainsLocatedEvidenceAndPreservesDestination(string extension) {
        string directory = CreateDirectory();
        try {
            string input = Path.Combine(directory, "source" + extension), destination = Path.Combine(directory, "result.pdf");
            CreateDrawing(input);
            byte[] originalInput = File.ReadAllBytes(input), originalOutput = [1, 2, 3];
            File.WriteAllBytes(destination, originalOutput);
            await using var output = new MemoryStream();
            using var error = new StringWriter();

            int exit = await OfficeImoToolApp.RunAsync(
                ["convert", input, destination, "--force", "--require-no-loss"], Stream.Null, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, exit);
            Assert.Contains("ODF_SKIPPED", error.ToString(), StringComparison.Ordinal);
            Assert.Contains("[document-metadata]", error.ToString(), StringComparison.Ordinal);
            Assert.Empty(output.ToArray());
            Assert.Equal(originalInput, File.ReadAllBytes(input));
            Assert.Equal(originalOutput, File.ReadAllBytes(destination));
            Assert.Empty(Directory.EnumerateFiles(directory, "*.tmp"));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(".odg", "--max-input-bytes", (int)OfficeImoToolExitCode.UnsupportedInput)]
    [InlineData(".fodg", "--max-input-bytes", (int)OfficeImoToolExitCode.UnsupportedInput)]
    [InlineData(".odg", "--max-output-bytes", (int)OfficeImoToolExitCode.OutputFailed)]
    [InlineData(".fodg", "--max-output-bytes", (int)OfficeImoToolExitCode.OutputFailed)]
    [InlineData(".odg", "--max-characters-in-part", (int)OfficeImoToolExitCode.UnsupportedInput)]
    [InlineData(".fodg", "--max-characters-in-part", (int)OfficeImoToolExitCode.UnsupportedInput)]
    public async Task DrawResourceLimitsPreserveExistingDestination(string extension, string option, int expected) {
        string directory = CreateDirectory();
        try {
            string input = Path.Combine(directory, "source" + extension), destination = Path.Combine(directory, "result.pdf");
            CreateDrawing(input);
            byte[] originalInput = File.ReadAllBytes(input), originalOutput = [1, 2, 3];
            File.WriteAllBytes(destination, originalOutput);
            await using var output = new MemoryStream();
            using var error = new StringWriter();

            int exit = await OfficeImoToolApp.RunAsync(
                ["convert", input, destination, "--force", option, "64"], Stream.Null, output, error);

            Assert.Equal(expected, exit);
            Assert.NotEmpty(error.ToString());
            Assert.Empty(output.ToArray());
            Assert.Equal(originalInput, File.ReadAllBytes(input));
            Assert.Equal(originalOutput, File.ReadAllBytes(destination));
            Assert.Empty(Directory.EnumerateFiles(directory, "*.tmp"));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(".odg")]
    [InlineData(".fodg")]
    public async Task DefaultDrawDestinationRequiresForceForReplacement(string extension) {
        string directory = CreateDirectory();
        try {
            string input = Path.Combine(directory, "source" + extension), destination = Path.ChangeExtension(input, ".pdf");
            CreateDrawing(input);
            byte[] original = [1, 2, 3];
            File.WriteAllBytes(destination, original);
            await using var output = new MemoryStream();
            using var error = new StringWriter();

            int refused = await OfficeImoToolApp.RunAsync(["convert", input], Stream.Null, output, error);
            Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, refused);
            Assert.Contains("--force", error.ToString(), StringComparison.Ordinal);
            Assert.Equal(original, File.ReadAllBytes(destination));
            Assert.Empty(output.ToArray());

            int replaced = await OfficeImoToolApp.RunAsync(["convert", input, "--force"], Stream.Null, output, error);
            Assert.Equal((int)OfficeImoToolExitCode.Success, replaced);
            Assert.Equal(2, PdfReadDocument.Open(File.ReadAllBytes(destination)).Pages.Count);
            Assert.Equal(destination + Environment.NewLine, Encoding.UTF8.GetString(output.ToArray()));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(".odg")]
    [InlineData(".fodg")]
    public async Task CancelledDrawConversionPreservesDestination(string extension) {
        string directory = CreateDirectory();
        try {
            string input = Path.Combine(directory, "source" + extension), destination = Path.Combine(directory, "result.pdf");
            CreateDrawing(input);
            byte[] original = [1, 2, 3];
            File.WriteAllBytes(destination, original);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await using var output = new MemoryStream();
            using var error = new StringWriter();

            int exit = await OfficeImoToolApp.RunAsync(
                ["convert", input, destination, "--force"], Stream.Null, output, error, cancellation.Token);

            Assert.Equal((int)OfficeImoToolExitCode.Cancelled, exit);
            Assert.Equal(original, File.ReadAllBytes(destination));
            Assert.Empty(output.ToArray());
            Assert.Empty(Directory.EnumerateFiles(directory, "*.tmp"));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("source.docx", "result.pdf", "--diagram-layers", "screen")]
    [InlineData("source.docx", "result.pdf", "--require-no-loss", null)]
    [InlineData("source.fodg", "result.md", "--diagram-layers", "screen")]
    [InlineData("source.fodg", "result.md", "--require-no-loss", null)]
    [InlineData("source.fodg", "result.pdf", "--diagram-layers", "unknown")]
    public async Task DiagramOptionsRejectUnrelatedRoutesAndInvalidLayerIntent(string input, string destination, string option, string? value) {
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        string[] args = value is null ? ["convert", input, destination, option] : ["convert", input, destination, option, value];

        int exit = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, exit);
        Assert.Empty(output.ToArray());
        Assert.NotEmpty(error.ToString());
    }

    private static string CreateDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static OdfRect Bounds(double y) => new(
        OdfLength.Points(20), OdfLength.Points(y), OdfLength.Points(250), OdfLength.Points(60));

    private static void CreateDrawing(string path) {
        var drawing = OdgDocument.Create();
        drawing.Metadata.Title = "Unprojected CLI metadata";
        drawing.AddPage("Main", OdfLength.Points(360), OdfLength.Points(220))
            .Shapes.AddTextBox(Bounds(20), "Draw CLI caption");
        drawing.AddPage("Blank", OdfLength.Points(180), OdfLength.Points(120));
        if (Path.GetExtension(path) == ".fodg") drawing.SaveFlatXml(path); else drawing.Save(path);
    }
}
