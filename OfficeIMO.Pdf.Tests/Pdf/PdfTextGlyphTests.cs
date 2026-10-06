using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTextGlyphTests {
    [Fact]
    public void ReaderExposesImmutableGlyphCodesAndSpacingOrigins() {
        PdfDocument document = PdfDocument.Create(pdf => pdf.Content(c => c.Paragraph(p => p.Text("Glyph"))));
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(document.ToBytes()).Pages[0].GetTextSpans());
        Assert.True(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs));
        Assert.Equal(span.Text, string.Concat(glyphs.Select(glyph => glyph.Text)));
        Assert.Equal(span.Text.Length, glyphs.Count);
        Assert.Equal(span.X, glyphs[0].X, 5);
        Assert.Equal(span.Y, glyphs[0].Y, 5);
        Assert.Equal(span.Advance, glyphs.Sum(glyph => glyph.Advance), 5);
        for (int index = 0; index < glyphs.Count; index++) {
            Assert.Equal(index, glyphs[index].TextStart);
            Assert.Equal(1, glyphs[index].TextLength);
            Assert.Equal((byte)span.Text[index], Assert.Single(glyphs[index].EncodedBytes));
        }
        Assert.Throws<NotSupportedException>(() => ((IList<byte>)glyphs[0].EncodedBytes)[0] = 0);
        Assert.Throws<NotSupportedException>(() => ((IList<PdfTextGlyph>)glyphs).Clear());
        Assert.True(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> repeated));
        Assert.Equal((byte)'G', Assert.Single(repeated[0].EncodedBytes));
    }

    [Fact]
    public void SyntheticSpanDoesNotInventGlyphEvidence() {
        var span = new PdfTextSpan("Logical only", "F1", 12, 10, 20);
        Assert.False(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs));
        Assert.Empty(glyphs);
    }
}
