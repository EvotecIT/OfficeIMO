using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRotatedFragmentExtractionTests {
    [Theory]
    [InlineData(0D)]
    [InlineData(5.729577951308233D)]
    [InlineData(-5.729577951308233D)]
    [InlineData(80D)]
    [InlineData(-80D)]
    public void PublicExtractionPreservesFragmentedRotatedLines(double angle) {
        byte[] bytes = CreatePdf(canvas => canvas.Effect(OfficeTransform.RotateDegrees(angle, 100D, 100D), 1D,
            rotated => {
                AddGlyphs(rotated, "LETTER_END", 100D, 100D);
                AddGlyphs(rotated, "NESTED_END", 100D, 125D);
            }));
        PdfReadPage page = PdfReadDocument.Open(bytes).Pages[0];

        string[] lines = page.ExtractText().Split('\n');
        Assert.Contains("LETTER_END", lines);
        Assert.Contains("NESTED_END", lines);
        Assert.Equal(2, lines.Length);
        string[] columnLines = page.ExtractTextWithColumns(new PdfTextLayoutOptions { ForceSingleColumn = true }).Split('\n');
        Assert.Contains("LETTER_END", columnLines);
        Assert.Contains("NESTED_END", columnLines);
        Assert.Equal(2, columnLines.Length);
    }

    [Theory]
    [InlineData(5.729577951308233D)]
    [InlineData(-5.729577951308233D)]
    public void MixedOrientationsDoNotInterleaveFragments(double angle) {
        byte[] bytes = CreatePdf(canvas => {
            AddGlyphs(canvas, "CAPTION", 250D, 100D);
            canvas.Effect(OfficeTransform.RotateDegrees(angle, 100D, 100D), 1D,
                rotated => AddGlyphs(rotated, "NESTED_END", 100D, 100D));
        });

        string[] lines = PdfReadDocument.Open(bytes).Pages[0].ExtractText().Split('\n');

        Assert.Equal(2, lines.Length);
        Assert.Contains("CAPTION", lines);
        Assert.Contains("NESTED_END", lines);
    }

    [Fact]
    public void LineGroupingPreservesOriginalSpanGeometry() {
        double angle = 0.1D, cos = Math.Cos(angle), sin = Math.Sin(angle);
        var spans = "NESTED_END".Select((glyph, index) => new PdfTextSpan(glyph.ToString(), "F1", 12D,
            50D + index * 7.2D * cos, 700D + index * 7.2D * sin, 7.2D,
            rotationDegrees: angle * 180D / Math.PI)).ToArray();
        var geometry = spans.Select(span => (span.X, span.Y, span.Advance, span.RotationDegrees)).ToArray();

        TextLayoutEngine.TextLine line = Assert.Single(TextLayoutEngine.BuildLines(spans));

        Assert.Equal("NESTED_END", line.Text);
        Assert.Equal(spans.Length, line.Spans.Count);
        Assert.All(line.Spans, span => Assert.Contains(span, spans));
        Assert.Equal(geometry, spans.Select(span => (span.X, span.Y, span.Advance, span.RotationDegrees)).ToArray());
    }

    [Theory]
    [InlineData(0D, false)]
    [InlineData(90D, false)]
    [InlineData(-90D, false)]
    [InlineData(90D, true)]
    public void DistinctRepeatedLabelsSurviveDuplicateLineDetection(double angle, bool fragments) {
        byte[] bytes = CreatePdf(canvas => canvas.Effect(OfficeTransform.RotateDegrees(angle, 100D, 100D), 1D,
            rotated => {
                foreach (double y in new[] { 100D, 125D }) {
                    if (fragments) AddGlyphs(rotated, "REPEATED_LABEL", 100D, y);
                    else rotated.Text("REPEATED_LABEL", 100D, y, 180D, 18D,
                        fontSize: 12D, font: PdfStandardFont.Courier);
                }
            }));
        PdfReadPage page = PdfReadDocument.Open(bytes).Pages[0];
        Assert.Equal(new[] { "REPEATED_LABEL", "REPEATED_LABEL" }, page.ExtractText().Split('\n'));
        Assert.Equal(new[] { "REPEATED_LABEL", "REPEATED_LABEL" }, page.ExtractTextWithColumns(
            new PdfTextLayoutOptions { ForceSingleColumn = true }).Split('\n'));
    }

    private static byte[] CreatePdf(Action<PdfPageCanvas> build) => PdfDocument.Create(new PdfOptions {
        PageWidth = 600D, PageHeight = 800D, CompressContentStreams = false
    }).Canvas(build).ToBytes();

    private static void AddGlyphs(PdfPageCanvas canvas, string text, double x, double y) {
        for (int index = 0; index < text.Length; index++) {
            string glyph = text[index].ToString();
            double glyphX = x + index * 7.2D;
            canvas.ActualText(glyph, painted => painted.Text(glyph, glyphX, y, 12D, 18D,
                fontSize: 12D, font: PdfStandardFont.Courier));
        }
    }
}
