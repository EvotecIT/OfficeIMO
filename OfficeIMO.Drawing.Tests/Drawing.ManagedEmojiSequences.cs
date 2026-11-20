using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

// UTS #51 sequences combine through the font's default ccmp lookups. Variation selectors are
// not ligature components, and no control glyph is painted after substitution.
public sealed class DrawingManagedEmojiSequencesTests {
    private const int BlankSpaceGlyph = 1;

    [Theory]
    [InlineData("\U0001F3CB\u200D\u2642", "\U0001F3CB\u200D\u2642", 0)]
    [InlineData("\U0001F3CB\uFE0F\u200D\u2642\uFE0F", "\U0001F3CB\u200D\u2642", 0)]
    [InlineData("\U0001F44D\U0001F3FB", "\U0001F44D\U0001F3FB", 1)]
    [InlineData("#\uFE0F\u20E3", "#\u20E3", 2)]
    [InlineData("\U0001F1F5\U0001F1F1", "\U0001F1F5\U0001F1F1", 3)]
    [InlineData("\U0001F3F4\U000E0067\U000E0062\U000E0065\U000E006E\U000E0067\U000E007F", "\U0001F3F4\U000E0067\U000E0062\U000E0065\U000E006E\U000E0067\U000E007F", 4)]
    [InlineData("\U0001F468\u200D\U0001F469", "\U0001F468\u200D\U0001F469", 5)]
    public void SequenceShapesToTheFontLigatureWithCompleteSource(string text, string unqualified, int ligature) {
        Assert.Equal(unqualified, ManagedTextShapingTestAssets.EmojiSequenceLigatures[ligature].Text);
        int ligatureGlyph = ManagedTextShapingTestAssets.EmojiSequenceLigatures[ligature].Ligature;

        OfficeTextShapingResult result = Shape(text);

        OfficeShapedGlyph painted = Assert.Single(result.Glyphs, glyph => glyph.GlyphId != BlankSpaceGlyph);
        Assert.Equal(ligatureGlyph, painted.GlyphId);
        Assert.Equal(text, painted.UnicodeText);
        Assert.Equal(ManagedTextShapingTestAssets.EmojiAdvance, result.Glyphs.Sum(glyph => glyph.AdvanceWidth));
        Assert.All(result.Glyphs.Where(glyph => glyph.GlyphId == BlankSpaceGlyph), glyph => Assert.Equal(0, glyph.AdvanceWidth));
    }

    [Theory]
    [InlineData("\U0001F44D\u200D\u2642\uFE0F", 2)]
    [InlineData("f\u200Di", 2)]
    [InlineData("f\u200Ci", 2)]
    [InlineData("\u2764\uFE0F", 1)]
    [InlineData("\U0001F3CB\u200C\u2642", 2)]
    public void UnjoinedSequenceKeepsComponentsAndNeverPaintsItsControls(string text, int components) {
        OfficeTextShapingResult result = Shape(text);

        OfficeShapedGlyph[] painted = result.Glyphs.Where(glyph => glyph.GlyphId != BlankSpaceGlyph).ToArray();
        Assert.Equal(components, painted.Length);
        Assert.All(painted, glyph => Assert.True(glyph.AdvanceWidth > 0));
        Assert.All(result.Glyphs.Where(glyph => glyph.GlyphId == BlankSpaceGlyph), glyph => Assert.Equal(0, glyph.AdvanceWidth));
        Assert.Equal(text, string.Concat(result.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
    }

    [Fact]
    public void VariationSequenceListedInCmapFormat14SelectsItsGlyph() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithNonDefaultVariationSequence(0x2764, 0xFE0F, blankSpace: true);

        OfficeTextShapingResult result = Shape("\u2764\uFE0F\u2764", data);

        Assert.Equal(new[] { 2, 3, 1 }, result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal(0, result.Glyphs[1].AdvanceWidth);
        Assert.Equal("\u2764\uFE0F", result.Glyphs[0].UnicodeText);
    }

    private static OfficeTextShapingResult Shape(string text, byte[]? data = null) {
        data ??= ManagedTextShapingTestAssets.CreateEmojiSequenceFont();
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Emoji sequences", data).Faces).Program;
        return Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
            text, "Emoji sequences", data, false, font.UnitsPerEm, OfficeTextDirection.LeftToRight, "en")));
    }
}
