using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.HarfBuzz;
using Xunit;

namespace OfficeIMO.Drawing.HarfBuzz.Tests;

public sealed class OfficeHarfBuzzTextShapingProviderTests {
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

    private static string GlyphSignature(OfficeTextShapingResult result) =>
        string.Join(
            "|",
            result.Glyphs.Select(static glyph =>
                $"{glyph.GlyphId}:{glyph.TextIndex}:{glyph.UnicodeText}:{glyph.AdvanceWidth}:{glyph.OffsetX}:{glyph.OffsetY}"));

    private static string FontPath(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "Fonts", fileName);
}
