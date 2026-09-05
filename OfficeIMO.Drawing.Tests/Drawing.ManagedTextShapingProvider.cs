using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingManagedTextShapingProviderTests {
    [Fact]
    public void ManagedProvider_ShapesSupportedArabicAndPreservesLogicalMappings() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(
            0x0627,
            0x0628,
            0xFE8D,
            0xFE8F);
        var request = new OfficeTextShapingRequest(
            "اب",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft,
            language: "ar");

        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);

        Assert.NotNull(result);
        Assert.Equal(2, result!.Glyphs.Count);
        Assert.Equal("ب", result.Glyphs[0].UnicodeText);
        Assert.Equal(1, result.Glyphs[0].TextIndex);
        Assert.Equal("ا", result.Glyphs[1].UnicodeText);
        Assert.Equal(0, result.Glyphs[1].TextIndex);
    }

    [Fact]
    public void ManagedProvider_DeclinesScriptsOutsideItsBoundedSubset() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(0x0915, 0x093F);
        var request = new OfficeTextShapingRequest(
            "कि",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000);

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
    }

    [Fact]
    public void ManagedProvider_PreservesMappingsWhenVisualGlyphsRepeat() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(0x0627, 0xFE8D);
        var request = new OfficeTextShapingRequest(
            "اا",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft,
            language: "ar");

        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);

        Assert.NotNull(result);
        Assert.Equal(new[] { 1, 0 }, result!.Glyphs.Select(static glyph => glyph.TextIndex));
    }

    [Fact]
    public void ManagedProvider_HonorsExplicitBaseDirectionForMixedText() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(
            ' ',
            'a',
            'b',
            'c',
            0x0627,
            0x0628,
            0xFE8D,
            0xFE8F);
        var request = new OfficeTextShapingRequest(
            "abc اب",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft,
            language: "ar");

        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);

        Assert.NotNull(result);
        Assert.Equal(new[] { 5, 4, 3, 0, 1, 2 }, result!.Glyphs.Select(static glyph => glyph.TextIndex));
    }

    [Fact]
    public void ManagedProvider_MapsExplicitBidiOverridesThroughSharedResolver() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(0x61, 0x62, 0x63);
        var request = new OfficeTextShapingRequest(
            "\u202Eabc\u202C",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft);

        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);

        Assert.NotNull(result);
        Assert.Equal(new[] { 3, 2, 1 }, result!.Glyphs.Select(static glyph => glyph.TextIndex));
    }

    [Fact]
    public void ManagedProvider_AppliesKerningToThePreviousVisualGlyphAdvance() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', adjustment: -120);
        var request = new OfficeTextShapingRequest(
            "\u202EVA\u202C",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft);

        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);

        Assert.NotNull(result);
        Assert.Equal(new[] { 1, 2 }, result!.Glyphs.Select(static glyph => glyph.GlyphId));
        Assert.All(result.Glyphs, static glyph => Assert.Null(glyph.AdvanceWidth));
        Assert.Equal(-120, result.GetAdvanceAdjustment(0));
        Assert.Equal(0, result.GetAdvanceAdjustment(1));
        OfficeTrueTypeFont loaded = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(font));
        Assert.Equal(880D, loaded.CreateShapedTextRun(request.Text, result).Measure(fontSize: 1000D), 6);
    }

    [Fact]
    public void ManagedProvider_HonorsExplicitKerningDisable() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', adjustment: -120);
        var request = new OfficeTextShapingRequest(
            "AV",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            featureSettings: new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("kern", 0) }),
            direction: OfficeTextDirection.LeftToRight,
            language: "en");

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(request));

        Assert.Equal(0, result.GetAdvanceAdjustment(0));
        Assert.Equal(0, result.GetAdvanceAdjustment(1));
    }

    [Fact]
    public void ManagedProvider_AppliesRequestedGsubLigatureAndPreservesExtractionText() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i');
        OfficeOpenTypeSubstitution substitution = Assert.IsType<OfficeOpenTypeSubstitution>(OfficeOpenTypeSubstitution.TryCreate(font));
        var tokens = new List<OfficeOpenTypeSubstitution.GlyphToken> {
            new OfficeOpenTypeSubstitution.GlyphToken(1, "f", 0, 'f'),
            new OfficeOpenTypeSubstitution.GlyphToken(2, "i", 1, 'i')
        };
        substitution.Apply(tokens, new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("liga", 1) }), default);
        Assert.Single(tokens);
        var request = new OfficeTextShapingRequest(
            "fi",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            featureSettings: new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("liga", 1) }),
            direction: OfficeTextDirection.LeftToRight,
            language: "en");

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(request));

        OfficeShapedGlyph glyph = Assert.Single(result.Glyphs);
        Assert.Equal(3, glyph.GlyphId);
        Assert.Equal("fi", glyph.UnicodeText);
        Assert.Equal(0, glyph.TextIndex);
    }
}
