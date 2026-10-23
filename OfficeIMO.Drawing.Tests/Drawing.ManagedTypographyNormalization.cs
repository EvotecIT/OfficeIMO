using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingManagedTypographyNormalizationTests {
    [Theory]
    [InlineData("لآَ")]
    [InlineData("لاَ\u0653")]
    public void ArabicMaddaCompositionUsesTheSelectedFontsLamalefAndVowelGeometry(string source) {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "NotoSansArabic-Regular.ttf"));
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Arabic typography", data).Faces).Program;

        OfficeTextShapingResult shaped = ShapeArabic(source, data, font, OfficeTextFeatureSettings.Default);

        // Exact reference-shaping contract for the retained Noto font: its required
        // lamalef-with-madda ligature has a different advance from plain lamalef.
        Assert.Equal(new[] { 292, 713 }, shaped.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal(662, shaped.Glyphs.Sum(glyph => glyph.AdvanceWidth));
        Assert.Equal(0, shaped.Glyphs[0].AdvanceWidth);
        Assert.Equal(39, shaped.Glyphs[0].OffsetX);
        Assert.Equal(192, shaped.Glyphs[0].OffsetY);
        AssertExactLogicalCoverage(source, shaped);
    }

    [Fact]
    public void ArabicMarkReorderingHonorsTheCombiningGraphemeJoinerOverride() {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "NotoSansArabic-Regular.ttf"));
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Arabic AMTRA", data).Faces).Program;
        Assert.True(font.TryGetGlyphMetrics(0x064E, out int fathaGlyph, out _));
        Assert.True(font.TryGetGlyphMetrics(0x0651, out int shaddaGlyph, out _));
        // Disabling ccmp keeps the separate mark glyphs observable instead of allowing
        // the font to combine shadda and a vowel after transient mark reordering.
        OfficeTextFeatureSettings features = OfficeTextFeatureSettings.Default.With("ccmp", 0);
        const string ordinarySource = "ب\u064E\u0651";
        const string overriddenSource = "ب\u064E\u034F\u0651";
        OfficeTextShapingResult ordinary = ShapeArabic(ordinarySource, data, font, features);
        OfficeTextShapingResult overridden = ShapeArabic(overriddenSource, data, font, features);

        // The result is in visual order. The observable contract is the selected
        // font's attachment geometry, independently confirmed against reference
        // shaping, rather than the write order of marks inside a visual cluster.
        OfficeShapedGlyph Mark(OfficeTextShapingResult shaped, int glyph) =>
            Assert.Single(shaped.Glyphs, item => item.GlyphId == glyph);
        Assert.Equal(40, Mark(ordinary, fathaGlyph).OffsetY);
        Assert.Equal(-160, Mark(ordinary, shaddaGlyph).OffsetY);
        Assert.Equal(-160, Mark(overridden, fathaGlyph).OffsetY);
        Assert.Equal(18, Mark(overridden, shaddaGlyph).OffsetY);
        AssertExactLogicalCoverage(ordinarySource, ordinary);
        AssertExactLogicalCoverage(overriddenSource, overridden);
    }

    [Theory]
    [InlineData('\u200C')]
    [InlineData('\u200D')]
    [InlineData('\u034F')]
    public void NonPaintingControlsKeepOriginalLogicalCoverageAndZeroAdvance(char control) {
        byte[] data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "NotoSansArabic-Regular.ttf"));
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Arabic AMTRA", data).Faces).Program;
        Assert.True(font.TryGetGlyphMetrics(' ', out int blankGlyph, out _));
        Assert.Empty(font.GetTextContours(" ", 0D, 0D, 16D));
        string source = "ل" + control + "ا";

        OfficeTextShapingResult shaped = ShapeArabic(source, data, font, OfficeTextFeatureSettings.Default);

        OfficeShapedGlyph ghost = Assert.Single(shaped.Glyphs, glyph => glyph.GlyphId == blankGlyph);
        Assert.Equal(0, ghost.AdvanceWidth);
        Assert.Equal(0, ghost.AdvanceHeight);
        AssertExactLogicalCoverage(source, shaped);
    }

    [Fact]
    public void CanonicalCompositionUsesCoveredGlyphWithoutChangingLogicalSource() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 0x00C1, 0x0301);
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Composition", data).Faces).Program;
        Assert.True(font.TryGetGlyphMetrics(0x00C1, out int composedGlyph, out _));

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest("A\u0301", "Composition", data, false, font.UnitsPerEm, OfficeTextDirection.LeftToRight, "en")));

        OfficeShapedGlyph glyph = Assert.Single(result.Glyphs);
        Assert.Equal(composedGlyph, glyph.GlyphId);
        Assert.Equal("A\u0301", glyph.UnicodeText);
        Assert.Equal(0, glyph.TextIndex);
    }

    [Fact]
    public void MissingPrecomposedGlyphUsesPositionedCanonicalComponentsWithSourceOwnedOnce() {
        byte[] data = ManagedTextShapingTestAssets.CreateMarkPositioningFont();
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Decomposition", data).Faces).Program;
        Assert.False(font.TryGetGlyphMetrics(0x00C1, out _, out _));
        const string source = "A\u0301\u0303";

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest(source, "Decomposition", data, false, font.UnitsPerEm, OfficeTextDirection.LeftToRight, "en")));

        Assert.Equal(new[] { 1, 2, 4 }, result.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(source, result.Glyphs[0].UnicodeText);
        Assert.All(result.Glyphs.Skip(1), glyph => { Assert.Empty(glyph.UnicodeText); Assert.Equal(0, glyph.AdvanceWidth); });
        Assert.Equal(source, string.Concat(result.Glyphs.Select(g => g.UnicodeText)));
        Assert.All(result.Glyphs, glyph => Assert.Equal(0, glyph.TextIndex));
    }

    [Fact]
    public void InvalidUtf16WithCombiningMarksDeclinesWithoutThrowing() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 0x00C1, 0x0301);
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Invalid source", data).Faces).Program;

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
            "A\u0301\uD800", "Invalid source", data, false, font.UnitsPerEm, OfficeTextDirection.LeftToRight, "en")));
    }

    [Theory]
    [InlineData("A\u0301\u035C")]
    [InlineData("A\u0301\uFFFF")]
    public void UnsupportedCharacterNeverReturnsAShapedPrefixOwningUnpaintedSource(string source) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 0x00C1, 0x0301);
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Missing source", data).Faces).Program;

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
            source, "Missing source", data, false, font.UnitsPerEm, OfficeTextDirection.LeftToRight, "en")));
    }

    private static OfficeTextShapingResult ShapeArabic(string source, byte[] data, IOfficeFontProgram font,
        OfficeTextFeatureSettings features) => Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest(source, "Arabic AMTRA", data, false, font.UnitsPerEm,
                OfficeTextDirection.RightToLeft, "ar", features)));

    private static void AssertExactLogicalCoverage(string source, OfficeTextShapingResult shaped) {
        var covered = new bool[source.Length];
        foreach (OfficeShapedGlyph glyph in shaped.Glyphs) {
            Assert.InRange(glyph.TextIndex, 0, source.Length - 1);
            Assert.InRange(glyph.UnicodeText.Length, 0, source.Length - glyph.TextIndex);
            Assert.Equal(glyph.UnicodeText, source.Substring(glyph.TextIndex, glyph.UnicodeText.Length));
            for (int index = glyph.TextIndex; index < glyph.TextIndex + glyph.UnicodeText.Length; index++) {
                Assert.False(covered[index], "A logical UTF-16 source position was owned by multiple glyphs.");
                covered[index] = true;
            }
        }
        Assert.All(covered, present => Assert.True(present, "A logical UTF-16 source position was omitted."));
    }
}
