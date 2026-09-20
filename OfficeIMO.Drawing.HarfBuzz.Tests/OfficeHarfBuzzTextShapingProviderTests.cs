using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.HarfBuzz.Tests;

public sealed class OfficeHarfBuzzTextShapingProviderTests {
    [Fact]
    public void ShapesTheSharedCrossRendererTypographyCorpus() {
        Assert.Equal(OfficeTextShapingBackend.HarfBuzz,
            ((IOfficeTextShapingProviderMetadata)OfficeHarfBuzzTextShapingProvider.Instance).Backend);

        foreach (TypographyEvidenceCase evidence in TypographyEvidenceCorpus.Cases) {
            byte[] fontData = LoadFontData(evidence);
            OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection()
                .Add(evidence.Family, fontData).Faces);
            Assert.True(face.Program.HasGlyphs(evidence.Text), evidence.Name);

            OfficeTextShapingResult? shaped = OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
                    evidence.Text,
                    evidence.Family,
                    face.Program.GetFontDataForShaping(),
                    face.Program.IsOpenTypeCff,
                    face.Program.UnitsPerEm,
                    evidence.Direction,
                    evidence.Language));
            Assert.Equal(evidence.HarfBuzzShapingExpected, shaped != null);
            if (shaped == null) continue;
            OfficeTextShapingResult result = shaped;

            Assert.Equal(evidence.Direction, result.Direction);
            Assert.NotEmpty(result.Glyphs);
            Assert.All(result.Glyphs, glyph => Assert.True(glyph.GlyphId > 0, evidence.Name));
            Assert.All(result.Glyphs, glyph =>
                Assert.Equal(glyph.UnicodeText,
                    evidence.Text.Substring(glyph.TextIndex, glyph.UnicodeText.Length)));
            if (evidence.Direction == OfficeTextDirection.TopToBottom) {
                Assert.All(result.Glyphs, glyph => Assert.NotEqual(0, glyph.AdvanceHeight));
                Assert.True(Math.Abs(result.Glyphs.Sum(glyph => glyph.AdvanceHeight ?? 0)) > 0);
            }
        }
    }

    [Fact]
    public void VerticalCjkUsesHarfBuzzAdvancesInRasterAndMarksSvgAsBrowserNative() {
        TypographyEvidenceCase evidence = Assert.Single(
            TypographyEvidenceCorpus.Cases,
            item => item.Direction == OfficeTextDirection.TopToBottom);
        byte[] fontData = LoadFontData(evidence);
        var drawing = new OfficeDrawing(120D, 180D)
            .AddFont(evidence.Family, fontData)
            .AddVerticalText(evidence.Text, 20D, 10D, 80D, 160D, new OfficeFontInfo(evidence.Family, 36D));

        var diagnostics = new List<OfficeImageExportDiagnostic>();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
            TextShapingProvider = OfficeHarfBuzzTextShapingProvider.Instance,
            TextShapingLanguage = evidence.Language,
            DiagnosticSink = diagnostics,
            DiagnosticSource = evidence.Name
        });
        (int width, int height) = InkSize(raster);

        Assert.True(height > width, $"Expected vertical ink, got {width}x{height}.");
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.TextShapingFallback);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        var text = Assert.Single(XDocument.Parse(svg).Descendants(), element => element.Name.LocalName == "text");
        Assert.Equal(evidence.Text, text.Value);
        Assert.Equal("vertical-rl", text.Attribute("writing-mode")?.Value);
        Assert.Equal("browser-native", text.Attribute("data-officeimo-shaping-backend")?.Value);
    }

    private static (int Width, int Height) InkSize(OfficeRasterImage image) {
        int minX = image.Width, minY = image.Height, maxX = -1, maxY = -1;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                if (image.GetPixel(x, y).A == 0) continue;
                minX = Math.Min(minX, x); minY = Math.Min(minY, y);
                maxX = Math.Max(maxX, x); maxY = Math.Max(maxY, y);
            }
        }
        Assert.True(maxX >= minX && maxY >= minY);
        return (maxX - minX + 1, maxY - minY + 1);
    }

    private static byte[] LoadFontData(TypographyEvidenceCase evidence) {
        if (!string.IsNullOrEmpty(evidence.FontFileName)) return File.ReadAllBytes(FontPath(evidence.FontFileName));
        return evidence.Name == "Hebrew"
            ? ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x05E9, 0x05DC, 0x05D5, 0x05DD, 0x05E2)
            : ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'C', 'a', 'f', 'e', 0x0301);
    }

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
            Assert.Same(first, current);
            Assert.Equal(expected, GlyphSignature(current));
        }
    }

    [Fact]
    public void CachedShapeResultsRespectUnitsPerEm() {
        const string text = "office";
        byte[] fontData = File.ReadAllBytes(FontPath("Carlito-Regular.ttf"));
        var fontCacheKey = new object();
        var smaller = new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 1024,
            OfficeTextDirection.LeftToRight,
            "en",
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            fontProgramCacheKey: fontCacheKey);
        var larger = new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            "en",
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            fontProgramCacheKey: fontCacheKey);

        OfficeTextShapingResult smallerResult = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(smaller));
        OfficeTextShapingResult largerResult = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(larger));

        Assert.NotSame(smallerResult, largerResult);
        int expectedLargerAdvance = smallerResult.Glyphs.Sum(static glyph => glyph.AdvanceWidth ?? 0) * 2;
        Assert.InRange(
            largerResult.Glyphs.Sum(static glyph => glyph.AdvanceWidth ?? 0),
            expectedLargerAdvance - 1,
            expectedLargerAdvance + 1);
    }

    [Fact]
    public void ExplicitOpenTypeFeaturesAreAppliedAndCachedSeparately() {
        const string text = "office";
        byte[] fontData = File.ReadAllBytes(FontPath("Carlito-Regular.ttf"));
        var fontCacheKey = new object();
        var defaultRequest = new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            "en",
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            fontProgramCacheKey: fontCacheKey);
        var withoutLigatures = new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            "en",
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            fontProgramCacheKey: fontCacheKey,
            featureSettings: OfficeTextFeatureSettings.Default.With("liga", 0));

        OfficeTextShapingResult defaultResult = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(defaultRequest));
        OfficeTextShapingResult disabledResult = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(withoutLigatures));
        OfficeTextShapingResult cachedDisabledResult = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(withoutLigatures));

        Assert.True(defaultResult.Glyphs.Count < text.Length);
        Assert.Equal(text.Length, disabledResult.Glyphs.Count);
        Assert.NotEqual(GlyphSignature(defaultResult), GlyphSignature(disabledResult));
        Assert.Same(disabledResult, cachedDisabledResult);
    }

    [Fact]
    public void OversizedLanguageHintsFallBackWithoutFragmentingTheCache() {
        const string text = "office";
        byte[] fontData = File.ReadAllBytes(FontPath("Carlito-Regular.ttf"));
        var provider = new OfficeHarfBuzzTextShapingProvider();
        var fontCacheKey = new object();
        var withoutLanguage = new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            language: null,
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            fontProgramCacheKey: fontCacheKey);
        var oversizedLanguage = new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            new string('a', 256),
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            fontProgramCacheKey: fontCacheKey);

        OfficeTextShapingResult first = Assert.IsType<OfficeTextShapingResult>(provider.ShapeText(withoutLanguage));
        OfficeTextShapingResult second = Assert.IsType<OfficeTextShapingResult>(provider.ShapeText(oversizedLanguage));

        Assert.Same(first, second);
    }

    [Fact]
    public void LanguageInterningIsNormalizedAndBoundedPerProvider() {
        const string text = "office";
        byte[] fontData = File.ReadAllBytes(FontPath("Carlito-Regular.ttf"));
        var provider = new OfficeHarfBuzzTextShapingProvider();
        var fontCacheKey = new object();
        OfficeTextShapingResult noLanguage = ShapeWithLanguage(provider, fontData, fontCacheKey, text, null);

        OfficeTextShapingResult normalized = ShapeWithLanguage(provider, fontData, fontCacheKey, text, " EN ");
        Assert.Same(normalized, ShapeWithLanguage(provider, fontData, fontCacheKey, text, "en"));

        for (int index = 1; index < OfficeHarfBuzzTextShapingProvider.MaxInternedLanguagesPerProvider; index++) {
            ShapeWithLanguage(provider, fontData, fontCacheKey, text, $"x-{index:x4}");
        }

        OfficeTextShapingResult overflow = ShapeWithLanguage(provider, fontData, fontCacheKey, text, "x-overflow");
        Assert.Same(noLanguage, overflow);
    }

    [Fact]
    public void OversizedShapeResultsAreNotRetained() {
        string text = new string('a', 4097);
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
        OfficeTextShapingResult second = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));

        Assert.NotSame(first, second);
        Assert.Equal(GlyphSignature(first), GlyphSignature(second));
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

    private static OfficeTextShapingResult ShapeWithLanguage(
        OfficeHarfBuzzTextShapingProvider provider,
        byte[] fontData,
        object fontCacheKey,
        string text,
        string? language) =>
        Assert.IsType<OfficeTextShapingResult>(provider.ShapeText(new OfficeTextShapingRequest(
            text,
            "Carlito",
            fontData,
            isOpenTypeCff: false,
            unitsPerEm: 2048,
            OfficeTextDirection.LeftToRight,
            language,
            default,
            fontCollectionIndex: null,
            variationCoordinates: null,
            cloneFontData: false,
            fontProgramCacheKey: fontCacheKey)));

    private static string FontPath(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "Fonts", fileName);
}
