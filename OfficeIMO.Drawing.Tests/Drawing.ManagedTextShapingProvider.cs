using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingManagedTextShapingProviderTests {
    [Fact]
    public void ManagedColorFont_ResolvesCpalLightAndDarkPalettes() {
        byte[] font = ManagedTextShapingTestAssets.CreateColorFont('A');
        OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection().Add("Color Test", font).Faces);
        IOfficeColorFontProgram program = Assert.IsAssignableFrom<IOfficeColorFontProgram>(face.Program);

        Assert.True(program.HasColorGlyph(1));
        Assert.True(program.TryGetColorLayers(1, "light", OfficeColor.Black, out IReadOnlyList<OfficeColorGlyphLayer> light));
        Assert.True(program.TryGetColorLayers(1, "dark", OfficeColor.Black, out IReadOnlyList<OfficeColorGlyphLayer> dark));
        Assert.Equal(new[] { OfficeColor.Red, OfficeColor.Blue }, light.Select(layer => layer.Color));
        Assert.Equal(new[] { OfficeColor.Yellow, OfficeColor.FromRgb(0, 128, 0) }, dark.Select(layer => layer.Color));
    }

    [Fact]
    public void ManagedColorFont_RasterizesSelectedPaletteLayers() {
        byte[] font = ManagedTextShapingTestAssets.CreateColorFont('A');
        var fonts = new OfficeFontFaceCollection().Add("Color Test", font);
        var lightImage = new OfficeRasterImage(80, 50, OfficeColor.White);
        var darkImage = new OfficeRasterImage(80, 50, OfficeColor.White);
        var lightCanvas = new OfficeRasterCanvas(lightImage, fonts: fonts);
        var darkCanvas = new OfficeRasterCanvas(darkImage, fonts: fonts);

        lightCanvas.DrawPositionedText("A", 2D, 2D, 60D, 44D, OfficeColor.Black, 36D, OfficeTextAlignment.Left,
            OfficeFontStyle.Regular, "Color Test", 24D, OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None,
            featureSettings: null, fontPalette: "light");
        darkCanvas.DrawPositionedText("A", 2D, 2D, 60D, 44D, OfficeColor.Black, 36D, OfficeTextAlignment.Left,
            OfficeFontStyle.Regular, "Color Test", 24D, OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None,
            featureSettings: null, fontPalette: "dark");

        Assert.True(ContainsColor(lightImage, pixel => pixel.R > 220 && pixel.G < 40 && pixel.B < 40));
        Assert.True(ContainsColor(lightImage, pixel => pixel.B > 220 && pixel.R < 40 && pixel.G < 40));
        Assert.True(ContainsColor(darkImage, pixel => pixel.R > 220 && pixel.G > 220 && pixel.B < 40));
        Assert.True(ContainsColor(darkImage, pixel => pixel.G > 90 && pixel.R < 40 && pixel.B < 40));
    }

    [Fact]
    public void VerticalColorFont_AppliesSimulatedBoldToEveryLayer() {
        byte[] font = ManagedTextShapingTestAssets.CreateColorFont('A');
        var fonts = new OfficeFontFaceCollection().Add("Vertical Color", font);
        var result = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 700, advanceHeight: -1000, offsetX: 0, offsetY: 0)
        }, OfficeTextDirection.TopToBottom);
        var provider = new FixedShapingProvider(result);
        var regularImage = new OfficeRasterImage(80, 80, OfficeColor.White);
        var boldImage = new OfficeRasterImage(80, 80, OfficeColor.White);
        var regular = new OfficeRasterCanvas(regularImage, font: null, fonts: fonts, textShapingProvider: provider);
        var bold = new OfficeRasterCanvas(boldImage, font: null, fonts: fonts, textShapingProvider: provider);

        Assert.True(regular.TryDrawVerticalText(
            "A", 10D, 10D, 50D, 60D, OfficeColor.Black, 36D,
            OfficeFontStyle.Regular, "Vertical Color", featureSettings: null, fontPalette: "light"));
        Assert.True(bold.TryDrawVerticalText(
            "A", 10D, 10D, 50D, 60D, OfficeColor.Black, 36D,
            OfficeFontStyle.Bold, "Vertical Color", featureSettings: null, fontPalette: "light"));

        Assert.True(CountInk(boldImage) > CountInk(regularImage));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShapedVerticalTextPaintsTypedDecorationsWithTheirOwnColor(bool colorFont) {
        byte[] font = colorFont
            ? ManagedTextShapingTestAssets.CreateColorFont('A')
            : ManagedTextShapingTestAssets.CreateFont('A');
        var fonts = new OfficeFontFaceCollection().Add("Vertical Decoration", font);
        var run = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 700, advanceHeight: -1000, offsetX: 0, offsetY: 0)
        }, OfficeTextDirection.TopToBottom);
        var image = new OfficeRasterImage(80, 80, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image, font: null, fonts: fonts, textShapingProvider: new FixedShapingProvider(run));

        Assert.True(canvas.TryDrawVerticalText(
            "A", 10D, 10D, 50D, 60D, OfficeColor.Black, 36D,
            OfficeFontStyle.Regular, "Vertical Decoration", featureSettings: null, fontPalette: "light",
            underlineStyle: OfficeTextDecorationStyle.Double,
            strikethroughStyle: OfficeTextDecorationStyle.Wavy,
            decorationColor: OfficeColor.FromRgb(255, 0, 255)));

        static bool IsDecoration(OfficeColor pixel) => pixel.R > 180 && pixel.G < 80 && pixel.B > 180;
        Assert.Contains(Enumerable.Range(30, 10), x => Enumerable.Range(10, 60).Any(y => IsDecoration(image.GetPixel(x, y))));
        Assert.Contains(Enumerable.Range(46, 12), x => Enumerable.Range(10, 60).Any(y => IsDecoration(image.GetPixel(x, y))));
    }

    [Fact]
    public void VerticalDrawingRetainsFontStyleDecorationsWithShapedAdvances() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        var run = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 700, advanceHeight: -1000, offsetX: 0, offsetY: 0)
        }, OfficeTextDirection.TopToBottom);
        var options = new OfficeDrawingRasterRenderOptions {
            TextShapingProvider = new FixedShapingProvider(run),
            Background = OfficeColor.White
        };
        static OfficeDrawing CreateDrawing(byte[] fontData, OfficeFontStyle style) =>
            new OfficeDrawing(80D, 80D)
                .AddFont("Vertical Decoration", fontData)
                .AddVerticalText("A", 10D, 10D, 50D, 60D,
                    new OfficeFontInfo("Vertical Decoration", 36D, style), OfficeColor.Black);

        OfficeRasterImage regular = OfficeDrawingRasterRenderer.Render(CreateDrawing(font, OfficeFontStyle.Regular), options);
        OfficeRasterImage decorated = OfficeDrawingRasterRenderer.Render(CreateDrawing(font,
            OfficeFontStyle.Underline | OfficeFontStyle.Strikethrough), options);

        Assert.True(CountInk(decorated) > CountInk(regular));
    }

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
    public void ManagedProvider_DeclinesTopToBottomRunsWithoutVerticalAdvances() {
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
            direction: OfficeTextDirection.TopToBottom,
            language: "ar");

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
    }

    private static int CountInk(OfficeRasterImage image) {
        byte[] pixels = image.GetPixels();
        int ink = 0;
        for (int i = 0; i + 3 < pixels.Length; i += 4) {
            if (pixels[i] < 250 || pixels[i + 1] < 250 || pixels[i + 2] < 250) ink++;
        }
        return ink;
    }

    private sealed class FixedShapingProvider : IOfficeTextShapingProvider {
        private readonly OfficeTextShapingResult _result;

        internal FixedShapingProvider(OfficeTextShapingResult result) {
            _result = result;
        }

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => _result;
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
    public void ManagedProvider_PreservesGposPlacementAndBothAdvanceAdjustments() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithPairPositioning('A', 'V');
        var request = new OfficeTextShapingRequest(
            "AV",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            featureSettings: new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("kern", 1) }),
            direction: OfficeTextDirection.LeftToRight,
            language: "en");

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(request));

        Assert.Equal(-10, result.Glyphs[0].OffsetX);
        Assert.Equal(-30, result.Glyphs[1].OffsetX);
        Assert.Equal(-20, result.GetAdvanceAdjustment(0));
        Assert.Equal(-40, result.GetAdvanceAdjustment(1));
        Assert.All(result.Glyphs, static glyph => Assert.Null(glyph.AdvanceWidth));
        OfficeTrueTypeFont loaded = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(font));
        Assert.Equal(940D, loaded.CreateShapedTextRun(request.Text, result).Measure(fontSize: 1000D), 6);
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

    [Fact]
    public void ManagedProvider_AppliesMultipleSubstitutionWithoutDuplicatingLogicalText() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithMultipleSubstitution('A');
        var request = new OfficeTextShapingRequest(
            "A",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.LeftToRight,
            language: "en",
            featureSettings: new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("ccmp", 1) }));

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(request));

        Assert.Equal(new[] { 3, 4 }, result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal("A", result.Glyphs[0].UnicodeText);
        Assert.Equal(string.Empty, result.Glyphs[1].UnicodeText);
        Assert.Equal(0, result.Glyphs[1].TextIndex);
    }

    [Fact]
    public void ManagedProvider_DoesNotReprocessGlyphsCreatedByMultipleSubstitution() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithSelfReferentialMultipleSubstitution('A');
        var request = new OfficeTextShapingRequest(
            "A",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.LeftToRight,
            language: "en",
            featureSettings: new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("ccmp", 1) }));

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(request));

        Assert.Equal(new[] { 2, 1 }, result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal(new[] { "A", string.Empty }, result.Glyphs.Select(glyph => glyph.UnicodeText));
    }

    [Fact]
    public void ManagedProvider_AppliesContextualFormatThreeLookupRecords() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithContextualSubstitution('A', 'B');
        var request = new OfficeTextShapingRequest(
            "AB",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.LeftToRight,
            language: "en",
            featureSettings: new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("calt", 1) }));

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(request));

        Assert.Equal(new[] { 1, 3 }, result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal(new[] { "A", "B" }, result.Glyphs.Select(glyph => glyph.UnicodeText));
    }

    [Fact]
    public void ManagedProvider_DeclinesUnsupportedFlagsOnNestedContextualLookups() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithUnsupportedNestedContextualLookupFlags('A', 'B');
        var request = new OfficeTextShapingRequest(
            "AB",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.LeftToRight,
            language: "en",
            featureSettings: new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("calt", 1) }));

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
    }

    [Fact]
    public void ManagedProvider_DeclinesNestedReverseContextualLookupsThatCannotRunAtOneGlyph() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithUnsupportedNestedReverseContextualLookup('A', 'B');
        var request = new OfficeTextShapingRequest(
            "AB",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.LeftToRight,
            language: "en",
            featureSettings: new OfficeTextFeatureSettings(new[] { new KeyValuePair<string, int>("calt", 1) }));

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
    }

    private static bool ContainsColor(OfficeRasterImage image, Func<OfficeColor, bool> predicate) {
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                if (predicate(image.GetPixel(x, y))) return true;
            }
        }
        return false;
    }
}
