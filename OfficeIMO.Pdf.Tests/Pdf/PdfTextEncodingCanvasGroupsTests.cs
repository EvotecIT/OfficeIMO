using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextEncodingCanvasGroupsTests {
    [Theory]
    [InlineData("structure")]
    [InlineData("actual-text")]
    [InlineData("artifact")]
    [InlineData("figure")]
    public void AnalyzeTextEncoding_TraversesNestedCanvasPaintWithoutTreatingLogicalUnicodeAsGlyphs(string group) {
        PdfDocument document = PdfDocument.Create().Canvas(canvas => {
            void Paint(PdfPageCanvas child) => child.Text("A−B", 0D, 0D, 100D, 20D);
            switch (group) {
                case "structure": canvas.Structure(PdfCanvasStructureRole.Paragraph, Paint); break;
                case "actual-text": canvas.ActualText("Logical Ω", Paint); break;
                case "artifact": canvas.Artifact(Paint); break;
                case "figure": canvas.Figure("Alternative Ω", Paint); break;
            }
        });

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(document.AnalyzeTextEncoding());

        Assert.Equal("U+2212", diagnostic.CodePoint);
        Assert.Equal("PdfCanvasText", diagnostic.Source);
        Assert.Contains("PdfCanvas", diagnostic.Location, StringComparison.Ordinal);
        Assert.Equal("unsupported-text-glyph", diagnostic.Code);
    }
}
