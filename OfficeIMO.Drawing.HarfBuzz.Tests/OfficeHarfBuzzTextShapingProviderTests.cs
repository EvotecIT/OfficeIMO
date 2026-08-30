using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.HarfBuzz;
using Xunit;

namespace OfficeIMO.Drawing.HarfBuzz.Tests;

public sealed class OfficeHarfBuzzTextShapingProviderTests {
    [Fact]
    public void RenderingProfileAppliesHarfBuzzToSharedExportOptions() {
        OfficeRenderingProfile profile = OfficeHarfBuzzRenderingProfile.Create(language: " ar ");
        var options = new OfficeImageExportOptions();

        options.UseRenderingProfile(profile);

        Assert.Equal("officeimo-harfbuzz", profile.Name);
        Assert.Same(OfficeHarfBuzzTextShapingProvider.Instance, options.TextShapingProvider);
        Assert.Equal("ar", options.TextShapingLanguage);
    }

    [Fact]
    public void ShapesLatinLigaturesWithLogicalClusterMappings() {
        const string text = "office";
        byte[] fontData = File.ReadAllBytes(FontPath("Carlito-Regular.ttf"));
        var request = new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            "en");

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));

        Assert.NotEmpty(result.Glyphs);
        Assert.True(result.Glyphs.Count < text.Length);
        Assert.All(result.Glyphs, glyph => {
            Assert.InRange(glyph.TextIndex, 0, text.Length - 1);
            Assert.Equal(
                glyph.UnicodeText,
                text.Substring(glyph.TextIndex, glyph.UnicodeText.Length));
        });
        Assert.Contains(result.Glyphs, glyph => glyph.UnicodeText.Length > 1);
    }

    [Fact]
    public void ShapesArabicWithPositionedVisualGlyphsAndLogicalText() {
        const string text = "سلام";
        byte[] fontData = File.ReadAllBytes(FontPath("NotoSansArabic-Regular.ttf"));
        var request = new OfficeTextShapingRequest(
            text,
            "Noto Sans Arabic",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            OfficeTextDirection.RightToLeft,
            "ar");

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));

        Assert.NotEmpty(result.Glyphs);
        Assert.All(result.Glyphs, glyph => {
            Assert.True(glyph.GlyphId > 0);
            Assert.NotEmpty(glyph.UnicodeText);
            Assert.InRange(glyph.TextIndex, 0, text.Length - 1);
        });
        Assert.Equal(
            text.OrderBy(static character => character),
            result.Glyphs.SelectMany(static glyph => glyph.UnicodeText).Distinct().OrderBy(static character => character));
    }

    [Fact]
    public void ShapesOpenTypeCffFontsThroughTheSameProviderContract() {
        const string text = "office";
        byte[] fontData = File.ReadAllBytes(FontPath("SourceSerif4-Regular.otf"));
        var request = new OfficeTextShapingRequest(
            text,
            "Source Serif 4",
            fontData,
            isOpenTypeCff: true,
            unitsPerEm: 1000,
            OfficeTextDirection.LeftToRight,
            "en");

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));

        Assert.NotEmpty(result.Glyphs);
        Assert.All(result.Glyphs, glyph => {
            Assert.True(glyph.GlyphId > 0);
            Assert.NotEmpty(glyph.UnicodeText);
            Assert.InRange(glyph.TextIndex, 0, text.Length - 1);
        });
    }

    [Theory]
    [InlineData("Noto Devanagari", "NotoSansDevanagari-Regular.ttf", "नमस्ते दुनिया", OfficeTextDirection.LeftToRight, "hi")]
    [InlineData("Noto CJK", "NotoSansSC-BaselineSubset.ttf", "永字国", OfficeTextDirection.LeftToRight, "zh")]
    [InlineData("Noto Emoji", "NotoEmoji-VariableFont_wght.ttf", "😀🚀🌍", OfficeTextDirection.LeftToRight, "und")]
    public void ShapesPortableScriptCorpusWithFirstPartyFontPrograms(
        string family,
        string fileName,
        string text,
        OfficeTextDirection direction,
        string language) {
        byte[] fontData = File.ReadAllBytes(FontPath(fileName));
        OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection().Add(family, fontData).Faces);
        Assert.True(face.Program.HasGlyphs(text));

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
                text,
                family,
                face.Program.GetFontDataForShaping(),
                face.Program.IsOpenTypeCff,
                face.Program.UnitsPerEm,
                direction,
                language)));

        Assert.NotEmpty(result.Glyphs);
        Assert.All(result.Glyphs, glyph => Assert.True(glyph.GlyphId > 0));
        Assert.NotEmpty(face.Program.GetTextContours(text, 0D, 0D, 24D));
    }

    [Fact]
    public void ReusesTheCachedNativeFontAcrossRepeatedShapes() {
        const string text = "office affinity efficient";
        byte[] fontData = File.ReadAllBytes(FontPath("Carlito-Regular.ttf"));
        var request = new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            "en");

        OfficeTextShapingResult first = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));
        string expected = GlyphSignature(first);

        for (int iteration = 0; iteration < 250; iteration++) {
            OfficeTextShapingResult current = Assert.IsType<OfficeTextShapingResult>(
                OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));
            Assert.Equal(expected, GlyphSignature(current));
        }
    }

    [Theory]
    [InlineData("Roboto Flex", "RobotoFlex.ttf", false, "A", "wght", 900F)]
    [InlineData("Adobe Variable CFF2", "AdobeVFPrototype-Subset.otf", true, "$", "wght", 700F)]
    public void ShapesTheSameSelectedVariableInstanceAsFirstPartyMetrics(
        string family,
        string fileName,
        bool isCff,
        string text,
        string axis,
        float value) {
        byte[] data = File.ReadAllBytes(FontPath(fileName));
        var fonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { [axis] = value }
        };
        fonts.Add(family, data);
        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.Equal(isCff, face.Program.IsOpenTypeCff);
        Assert.True(face.Program.TryGetGlyphMetrics(text[0], out int expectedGlyph, out int expectedAdvance));

        var coordinates = new Dictionary<string, float> { [axis] = value };
        var request = new OfficeTextShapingRequest(
            text,
            family,
            data,
            isCff,
            face.Program.UnitsPerEm,
            OfficeTextDirection.LeftToRight,
            "en",
            default,
            fontCollectionIndex: null,
            coordinates);
        coordinates[axis] = value == 700F ? 100F : 400F;

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));

        OfficeShapedGlyph glyph = Assert.Single(result.Glyphs);
        Assert.Equal(expectedGlyph, glyph.GlyphId);
        Assert.Equal(expectedAdvance, glyph.AdvanceWidth);
        Assert.Equal(value, request.VariationCoordinates[axis]);
    }

    private static string GlyphSignature(OfficeTextShapingResult result) =>
        string.Join(
            "|",
            result.Glyphs.Select(static glyph =>
                $"{glyph.GlyphId}:{glyph.TextIndex}:{glyph.UnicodeText}:{glyph.AdvanceWidth}:{glyph.OffsetX}:{glyph.OffsetY}"));

    private static string FontPath(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "Fonts", fileName);
}
