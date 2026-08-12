using System.Collections.Generic;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextLayoutDuplicateGlyphTests {
    [Fact]
    public void BuildLines_PreservesAdjacentRepeatedGlyphs() {
        List<PdfTextSpan> spans = CreateGlyphSpans("OfficeIMO");

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);

        Assert.Single(lines);
        Assert.Equal("OfficeIMO", lines[0].Text);
    }

    [Fact]
    public void BuildLines_RemovesSubstantiallyOverlappingShadowGlyph() {
        var spans = new List<PdfTextSpan> {
            new("A", "F1", 12, 10, 100, 6),
            new("A", "F1", 12, 10.2, 100, 6)
        };

        List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);

        Assert.Single(lines);
        Assert.Equal("A", lines[0].Text);
    }

    private static List<PdfTextSpan> CreateGlyphSpans(string text) {
        var spans = new List<PdfTextSpan>(text.Length);
        double x = 10;
        foreach (char glyph in text) {
            spans.Add(new PdfTextSpan(glyph.ToString(), "F1", 12, x, 100, 6));
            x += 6;
        }

        return spans;
    }
}
