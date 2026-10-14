using System.Collections.Generic;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextLayoutDuplicateGlyphTests {
    [Theory]
    [InlineData(0D)]
    [InlineData(5.729577951308233D)]
    [InlineData(-5.729577951308233D)]
    [InlineData(80D)]
    public void BuildLines_PreservesAdjacentRepeatedGlyphs(double angle) {
        List<PdfTextSpan> spans = CreateGlyphSpans("OfficeIMO", angle);

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);

        Assert.Single(lines);
        Assert.Equal("OfficeIMO", lines[0].Text);
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(5.729577951308233D)]
    [InlineData(-5.729577951308233D)]
    [InlineData(80D)]
    public void BuildLines_RemovesSubstantiallyOverlappingShadowGlyph(double angle) {
        double radians = angle * Math.PI / 180D;
        var spans = new List<PdfTextSpan> {
            new("A", "F1", 12, 10, 100, 6, rotationDegrees: angle),
            new("A", "F1", 12, 10 + 0.2D * Math.Cos(radians), 100 + 0.2D * Math.Sin(radians), 6,
                rotationDegrees: angle)
        };

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);

        Assert.Single(lines);
        Assert.Equal("A", lines[0].Text);
    }

    [Fact]
    public void DifferentBaselineDirectionsAreNotDuplicateLines() {
        var spans = new[] {
            new PdfTextSpan("Caption", "F1", 12D, 50D, 700D, 42D),
            new PdfTextSpan("Caption", "F1", 12D, 50D, 700D, 42D, rotationDegrees: 5.729577951308233D)
        };

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);

        Assert.Equal(2, lines.Count);
        Assert.All(lines, line => Assert.Equal("Caption", line.Text));
    }

    private static List<PdfTextSpan> CreateGlyphSpans(string text, double angle) {
        var spans = new List<PdfTextSpan>(text.Length);
        double radians = angle * Math.PI / 180D;
        double distance = 0D;
        foreach (char glyph in text) {
            spans.Add(new PdfTextSpan(glyph.ToString(), "F1", 12, 10 + distance * Math.Cos(radians),
                100 + distance * Math.Sin(radians), 6, rotationDegrees: angle));
            distance += 6;
        }

        return spans;
    }
}
