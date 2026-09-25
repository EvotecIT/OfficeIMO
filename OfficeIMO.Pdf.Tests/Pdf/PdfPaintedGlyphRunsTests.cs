using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPaintedGlyphRunsTests {
    [Fact]
    public void UnresolvedGlyphInSubstituteFontKeepsNeighboursAtPaintedPositions() {
        var spans = new List<PdfTextSpan> {
            new("A" + PdfPaintedGlyphRuns.UndecodedGlyph + "B", "F1", 12D, 10D, 30D, 30D,
                color: null, isVisible: true, rotationDegrees: 0D, baseFont: null, clipPath: null,
                characterAdvances: new[] { 10D, 10D, 10D },
                glyphCharacterLengths: new[] { 1, 1, 1 },
                glyphBytes: new byte[][] { [65], [0], [66] },
                glyphPaintedAdvances: new[] { 10D, 10D, 10D },
                characterAdvanceDirection: 1D)
        };
        int charged = 0;

        PdfPaintedGlyphRuns.SplitComplexRuns(spans, count => charged += count);

        Assert.Equal(3, charged);
        Assert.Equal(new[] { "A", PdfPaintedGlyphRuns.UndecodedGlyphText, "B" }, spans.Select(span => span.Text));
        Assert.Equal(new[] { 10D, 20D, 30D }, spans.Select(span => span.X));
    }
}
