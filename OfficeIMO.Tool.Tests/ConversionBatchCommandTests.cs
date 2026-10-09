using OfficeIMO.OpenDocument;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class ConversionBatchCommandTests {
    [Fact]
    public async Task BatchCommandExecutesAndResumesTheSharedContract() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-batch-cli-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(root, "input"));
        try {
            var request = new OfficeConversionBatchRequest {
                InputDirectory = Path.Combine(root, "input"), OutputDirectory = Path.Combine(root, "output"),
                CheckpointDirectory = Path.Combine(root, "state")
            };
            File.WriteAllText(Path.Combine(request.InputDirectory, "source.txt"), "Literal <b>content</b>");
            for (int run = 0; run < 2; run++) {
                using var output = new MemoryStream();
                using var error = new StringWriter();
                int code = await OfficeImoToolApp.RunAsync(["workflow", "batch", "--input-directory", request.InputDirectory!, "--output", request.OutputDirectory, "--checkpoint", request.CheckpointDirectory!], Stream.Null, output, error);
                Assert.Equal((int)OfficeImoToolExitCode.Success, code);
                string json = Encoding.UTF8.GetString(output.ToArray());
                Assert.Contains("\"Completed\":1", json);
                Assert.Contains("\"Reused\":" + run, json);
                Assert.Equal(string.Empty, error.ToString());
            }
            Assert.True(File.Exists(Path.Combine(request.OutputDirectory, "source.txt.pdf")));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task BatchDiagramAcceptanceDoesNotPreventLiteralTextConversion(bool strictLoss) {
        string root = CreateRoot();
        try {
            string input = Path.Combine(root, "input"), destination = Path.Combine(root, "output");
            CreateDraw(Path.Combine(input, "drawing.fodg"));
            File.WriteAllText(Path.Combine(input, "text.txt"), "Literal mixed batch marker");
            string[] acceptance = strictLoss ? ["--require-no-loss"] : ["--maximum-xml-characters-in-part", "256"];
            var run = await RunBatch(["--input-directory", input, "--output", destination, .. acceptance]);

            Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, run.Code);
            using JsonDocument result = JsonDocument.Parse(run.Output);
            Assert.Equal(1, result.RootElement.GetProperty("Completed").GetInt32());
            Assert.Equal(1, result.RootElement.GetProperty("Failed").GetInt32());
            Assert.Contains("Literal mixed batch marker",
                PdfReadDocument.Open(File.ReadAllBytes(Path.Combine(destination, "text.txt.pdf"))).ExtractText(), StringComparison.Ordinal);
            Assert.False(File.Exists(Path.Combine(destination, "drawing.fodg.pdf")));
            Assert.Contains("drawing.fodg", run.Error, StringComparison.Ordinal);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task BatchDiagramLayerOptionSelectsScreenContent() {
        string root = CreateRoot();
        try {
            string input = Path.Combine(root, "input"), destination = Path.Combine(root, "output");
            var drawing = OdgDocument.Create();
            var page = drawing.AddPage("Layers", OdfLength.Points(360), OdfLength.Points(220));
            page.Layers.Add("Screen", OdgLayerDisplay.Screen);
            page.Layers.Add("Print", OdgLayerDisplay.Printer);
            page.Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 1), "BatchScreenOnly").Layer = "Screen";
            page.Shapes.AddTextBox(OdfRect.FromCentimeters(1, 3, 8, 1), "BatchPrintOnly").Layer = "Print";
            drawing.SaveFlatXml(Path.Combine(input, "layers.fodg"));

            var run = await RunBatch(["--input-directory", input, "--output", destination, "--diagram-layers", "screen"]);

            Assert.Equal((int)OfficeImoToolExitCode.Success, run.Code);
            string text = PdfReadDocument.Open(File.ReadAllBytes(Path.Combine(destination, "layers.fodg.pdf"))).ExtractText();
            Assert.Contains("BatchScreenOnly", text, StringComparison.Ordinal);
            Assert.DoesNotContain("BatchPrintOnly", text, StringComparison.Ordinal);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ReusedBatchCheckpointRetainsPageQualifiedLossOnStandardError() {
        string root = CreateRoot();
        try {
            string input = Path.Combine(root, "input"), destination = Path.Combine(root, "output");
            CreateDraw(Path.Combine(input, "drawing.fodg"));
            string[] args = ["--input-directory", input, "--output", destination, "--checkpoint", Path.Combine(root, "state")];
            var first = await RunBatch(args);
            Assert.Equal((int)OfficeImoToolExitCode.Success, first.Code);
            string finding = first.Error.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
                .First(line => line.Contains("[page:1:", StringComparison.Ordinal) && line.Contains("capitalization", StringComparison.Ordinal));
            byte[] original = File.ReadAllBytes(Path.Combine(destination, "drawing.fodg.pdf"));

            var reused = await RunBatch(args);

            Assert.Equal((int)OfficeImoToolExitCode.Success, reused.Code);
            using JsonDocument result = JsonDocument.Parse(reused.Output);
            Assert.Equal(1, result.RootElement.GetProperty("Reused").GetInt32());
            Assert.Contains(finding, reused.Error, StringComparison.Ordinal);
            Assert.Equal(original, File.ReadAllBytes(Path.Combine(destination, "drawing.fodg.pdf")));
        } finally { Directory.Delete(root, true); }
    }

    private static string CreateRoot() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-draw-batch-cli-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(root, "input"));
        return root;
    }

    private static void CreateDraw(string path) {
        var drawing = OdgDocument.Create();
        drawing.AddPage("Native capitalization").Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 5), string.Empty)
            .AddParagraph("native capitalization").TextTransform = OdfTextTransform.Capitalize;
        drawing.SaveFlatXml(path);
    }

    private static async Task<(int Code, string Output, string Error)> RunBatch(string[] args) {
        using var output = new MemoryStream();
        using var error = new StringWriter();
        int code = await OfficeImoToolApp.RunAsync(["workflow", "batch", .. args], Stream.Null, output, error);
        return (code, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }
}
