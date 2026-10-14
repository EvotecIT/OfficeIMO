using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class DrawLossEvidenceCommandTests {
    [Fact]
    public async Task StrictDrawConversionRetainsLaterPageAndPdfSaveFailureDiagnostics() {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-draw-loss-evidence-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "source.fodg"), destination = Path.Combine(root, "result.pdf");
            var drawing = OdgDocument.Create();
            drawing.Metadata.Title = "Unprojected source metadata";
            drawing.AddPage("First").Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 5), string.Empty)
                .AddParagraph("native capitalization").TextTransform = OdfTextTransform.Capitalize;
            drawing.AddPage("Second").Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 5), "مرحبا");
            drawing.SaveFlatXml(input);
            byte[] originalSource = File.ReadAllBytes(input), original = { 1, 2, 3 };
            File.WriteAllBytes(destination, original);
            using var output = new MemoryStream();
            using var error = new StringWriter();
            int result = await OfficeImoToolApp.RunAsync(["convert", input, destination, "--force", "--require-no-loss"], Stream.Null, output, error);
            Assert.NotEqual((int)OfficeImoToolExitCode.Success, result);
            Assert.Contains("page:1:First", error.ToString(), StringComparison.Ordinal);
            Assert.Contains("page:2:Second", error.ToString(), StringComparison.Ordinal);
            Assert.True(error.ToString().Contains("unsupported-text-glyph", StringComparison.Ordinal), error.ToString());
            Assert.Empty(output.ToArray());
            Assert.Equal(original, File.ReadAllBytes(destination));
            Assert.Equal(originalSource, File.ReadAllBytes(input));
        } finally { Directory.Delete(root, true); }
    }
}
