using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextEditorTests {
    [Fact]
    public void GlyphAccessKeepsLigatureAndSupplementaryUnicodeRanges() {
        const string cmap = "/CIDInit /ProcSet findresource begin 12 dict begin begincmap " +
            "/CMapType 2 def 1 begincodespacerange <00> <FF> endcodespacerange " +
            "2 beginbfchar <41> <006600660069> <42> <D83DDE00> endbfchar endcmap end end\n";
        byte[] source = BuildRawTextPdf("BT /F1 12 Tf 50 700 Td (AB) Tj ET\n",
            additionalObjects: "7 0 obj\n<< /Length " + cmap.Length + " >>\nstream\n" + cmap + "endstream\nendobj\n");
        source = PdfEncoding.Latin1GetBytes(PdfEncoding.Latin1GetString(source).Replace(
            "/BaseFont /Helvetica >>", "/BaseFont /Helvetica /ToUnicode 7 0 R >>"));
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

        Assert.Equal("ffi\U0001F600", span.Text);
        Assert.True(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs));
        Assert.Collection(glyphs,
            glyph => { Assert.Equal("ffi", glyph.Text); Assert.Equal(0, glyph.TextStart); Assert.Equal(3, glyph.TextLength); Assert.Equal(new byte[] { 65 }, glyph.EncodedBytes); },
            glyph => { Assert.Equal("\U0001F600", glyph.Text); Assert.Equal(3, glyph.TextStart); Assert.Equal(2, glyph.TextLength); Assert.Equal(new byte[] { 66 }, glyph.EncodedBytes); });
        Assert.Equal(span.Advance, glyphs.Sum(glyph => glyph.Advance), precision: 6);
    }

    [Fact]
    public void GlyphAccessReportsAmbiguousActualTextWithoutInventingRanges() {
        byte[] source = BuildRawTextPdf("/Span << /ActualText (ffi) >> BDC BT /F1 12 Tf 50 700 Td (x) Tj ET EMC\n");
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());
        Assert.Equal("ffi", span.Text);
        Assert.False(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs));
        Assert.Empty(glyphs);
    }
}
