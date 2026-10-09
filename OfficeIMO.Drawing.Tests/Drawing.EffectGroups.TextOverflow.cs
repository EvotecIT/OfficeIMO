using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingEffectParagraphOverflowTests {
    private const string Family = "Effect Caption Proof";
    private static byte[] Font() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
    private static OfficeDrawing Caption(string value, OfficeTextAlignment alignment, bool wrap = false, double height = 40) =>
        new OfficeDrawing(20, height).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(value, 10, OfficeColor.Black, fontFamily: Family) }, alignment)
        }, 0, 0, 20, height, wrapText: wrap);

    [Theory]
    [InlineData(OfficeTextAlignment.Center, false)]
    [InlineData(OfficeTextAlignment.Right, false)]
    [InlineData(OfficeTextAlignment.Center, true)]
    [InlineData(OfficeTextAlignment.Right, true)]
    public void TranslatedEffectMatchesTheSameLogicalFrameWithLateFonts(OfficeTextAlignment alignment, bool fitting) {
        var inner = Caption(fitting ? "Fit" : "OVERWIDE CAPTION TEXT", alignment);
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(inner, OfficeTransform.Translate(140, 40));
        var expected = new OfficeDrawing(300, 160).AddDrawing(inner, 140, 40);
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(expected);
        Assert.Equal(!fitting, HasInkOutside(reference, 140, 160));
        EqualRaster(reference, OfficeDrawingRasterRenderer.Render(actual));
        Assert.Empty(inner.Fonts.Faces);
        Assert.Empty(Assert.IsType<OfficeDrawingEffectGroup>(actual.Elements[0]).Drawing.Fonts.Faces);
        var frame = Assert.IsType<OfficeDrawingRichText>(inner.Elements[0]);
        Assert.Equal((0D, 0D, 20D, 40D), (frame.X, frame.Y, frame.Width, frame.Height));
    }

    [Theory]
    [InlineData(1D, 2D)]
    [InlineData(2D, 2D)]
    [InlineData(1D, .5D)]
    [InlineData(2D, .5D)]
    public void NestedAffineEffectsMatchAnUncroppedContainerWithTheSameTextFrame(double scale, double horizontalScale) {
        var caption = Caption("OVERWIDE CAPTION TEXT", OfficeTextAlignment.Center);
        OfficeTransform childTransform = OfficeTransform.Scale(horizontalScale, 1).Then(OfficeTransform.Translate(5, 0));
        var narrow = new OfficeDrawing(45, 40).AddEffectDrawing(caption, childTransform);
        var wide = new OfficeDrawing(250, 40).AddEffectDrawing(caption, childTransform);
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(narrow, OfficeTransform.Translate(135, 40));
        var expected = new OfficeDrawing(300, 160).AddEffectDrawing(wide, OfficeTransform.Translate(135, 40));
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(expected, scale);
        Assert.True(HasInkOutside(reference, (int)(140 * scale), (int)((140 + 20 * horizontalScale) * scale)));
        EqualRaster(reference, OfficeDrawingRasterRenderer.Render(actual, scale));
    }

    [Fact]
    public void NestedEffectsRetainComposedTranslationAndInk() {
        var caption = Caption("OVERWIDE CAPTION TEXT", OfficeTextAlignment.Right);
        var nested = new OfficeDrawing(25, 40).AddEffectDrawing(caption, OfficeTransform.Translate(5, 0));
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(nested, OfficeTransform.Translate(135, 40));
        var expected = new OfficeDrawing(300, 160).AddDrawing(caption, 140, 40);
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());
        EqualRaster(OfficeDrawingRasterRenderer.Render(expected), OfficeDrawingRasterRenderer.Render(actual));
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(2D)]
    public void NestedReflectedCaptionRetainsShiftedMaskAndPhysicalAllocationLimit(double scale) {
        var caption = Caption("OVERWIDE CAPTION TEXT", OfficeTextAlignment.Center);
        var mask = new OfficeDrawing(60, 40);
        OfficeShape cover = OfficeShape.Rectangle(60, 40);
        cover.FillColor = OfficeColor.White; cover.StrokeWidth = 0;
        mask.AddShape(cover, 0, 0);
        var softMask = new OfficeDrawingSoftMask(mask, OfficeSoftMaskMode.Alpha, OfficeTransform.Translate(-30, 0));
        OfficeTransform child = OfficeTransform.Scale(.5, 2).Then(OfficeTransform.Translate(5, 0));
        var narrow = new OfficeDrawing(45, 80).AddEffectDrawing(caption, child, OfficeBlendMode.Normal, softMask);
        var wide = new OfficeDrawing(250, 80).AddEffectDrawing(caption, child, OfficeBlendMode.Normal, softMask);
        OfficeTransform parent = OfficeTransform.Scale(-2, .5).Then(OfficeTransform.Translate(200, 40));
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(narrow, parent);
        var expected = new OfficeDrawing(300, 160).AddEffectDrawing(wide, parent);
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());

        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(expected, scale);
        Assert.True(HasInkOutside(reference, (int)(170 * scale), (int)(190 * scale)));
        EqualRaster(reference, OfficeDrawingRasterRenderer.Render(actual,
            new OfficeDrawingRasterRenderOptions { Scale = scale, MaximumRasterPixels = 500_000 }));
        // The output ceiling and shared intermediate ceiling are independent.
        // A small visible viewport admits the output but rejects these expanded nested layers.
        var bounded = new OfficeDrawing(70, 30).AddEffectDrawing(narrow,
            parent.Then(OfficeTransform.Translate(-165, -40)));
        bounded.Fonts.Add(Family, Font());
        Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(bounded,
            new OfficeDrawingRasterRenderOptions { Scale = scale, MaximumRasterPixels = (long)(70 * 30 * scale * scale) }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InsertedFontAndLateParentReplacementUseTheEffectiveProfileWithoutMutation(bool replaceLate) {
        byte[] localBytes = Font(), parentBytes = ManagedTextShapingTestAssets.CreateFont('A');
        var inner = Caption("AAAA", OfficeTextAlignment.Center);
        inner.Fonts.Add(Family, localBytes);
        var actual = new OfficeDrawing(300, 160);
        actual.Fonts.Add(Family, parentBytes);
        actual.AddEffectDrawing(inner, OfficeTransform.Translate(140, 40));
        if (replaceLate) actual.Fonts.Add(Family, parentBytes);
        byte[] effective = replaceLate ? parentBytes : localBytes;
        var expected = new OfficeDrawing(300, 160).AddDrawing(Caption("AAAA", OfficeTextAlignment.Center), 140, 40);
        expected.Fonts.Add(Family, effective);
        EqualRaster(OfficeDrawingRasterRenderer.Render(expected), OfficeDrawingRasterRenderer.Render(actual));
        Assert.Equal(localBytes, Assert.Single(inner.Fonts.Faces).Data);
        Assert.Equal(localBytes, Assert.Single(Assert.IsType<OfficeDrawingEffectGroup>(actual.Elements[0]).Drawing.Fonts.Faces).Data);
        Assert.Equal(effective, Assert.Single(actual.Fonts.Faces).Data);
    }

    [Fact]
    public void MeasurementAndLayerPaintKeepTheChildShapingProfile() {
        var childProvider = new CaptionGlyphProvider(1400);
        var parentProvider = new CaptionGlyphProvider(400);
        var inner = Caption("AAAA", OfficeTextAlignment.Center);
        inner.Fonts.Add(Family, ManagedTextShapingTestAssets.CreateFont('A'));
        inner.ApplyImageExportOptions(new OfficeImageExportOptions { TextShapingProvider = childProvider, TextShapingLanguage = "ar-SA" });
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(inner, OfficeTransform.Translate(140, 40));
        actual.ApplyImageExportOptions(new OfficeImageExportOptions { TextShapingProvider = parentProvider, TextShapingLanguage = "en-US" });
        var expected = new OfficeDrawing(300, 160).AddDrawing(inner, 140, 40);
        expected.ApplyImageExportOptions(new OfficeImageExportOptions { TextShapingProvider = childProvider, TextShapingLanguage = "ar-SA" });
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(expected);
        Assert.True(HasInkOutside(reference, 140, 160));
        childProvider.Languages.Clear();
        EqualRaster(reference, OfficeDrawingRasterRenderer.Render(actual));
        Assert.NotEmpty(childProvider.Languages);
        Assert.All(childProvider.Languages, value => Assert.Equal("ar-SA", value));
        Assert.Empty(parentProvider.Languages);
    }

    [Fact]
    public void ExplicitClipAndVerticalSurfaceRemainBoundaries() {
        var inner = Caption("OVERWIDE CAPTION TEXT\nSECOND LINE\nTHIRD LINE", OfficeTextAlignment.Center, height: 15);
        var local = new OfficeDrawing(20, 15).AddClippedDrawing(inner, 0, 0, OfficeClipPath.Rectangle(20, 15));
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(local, OfficeTransform.Translate(140, 40));
        var expected = new OfficeDrawing(300, 160).AddClippedDrawing(inner, 140, 40, OfficeClipPath.Rectangle(20, 15));
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(actual);
        EqualRaster(OfficeDrawingRasterRenderer.Render(expected), raster);
        Assert.False(HasInkOutside(raster, 140, 160));
        Assert.All(Enumerable.Range(55, raster.Height - 55), y => Assert.All(Enumerable.Range(0, raster.Width), x => Assert.Equal((byte)0, raster.GetPixel(x, y).A)));
    }

    [Fact]
    public void WrappedParagraphsKeepTheirFrameAndEffectBudget() {
        var inner = Caption("OVERWIDE CAPTION TEXT", OfficeTextAlignment.Center, wrap: true);
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(inner, OfficeTransform.Translate(140, 40));
        var expected = new OfficeDrawing(300, 160).AddDrawing(inner, 140, 40);
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());
        EqualRaster(OfficeDrawingRasterRenderer.Render(expected), OfficeDrawingRasterRenderer.Render(actual));
        Assert.False(HasInkOutside(OfficeDrawingRasterRenderer.Render(actual), 140, 160));
        var budgetScene = new OfficeDrawing(20, 40).AddEffectDrawing(inner, OfficeTransform.Identity);
        budgetScene.Fonts.Add(Family, Font());
        Assert.NotNull(OfficeDrawingRasterRenderer.Render(budgetScene, new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 800 }));
    }

    [Fact]
    public void ExpandedInkIsChargedToTheIntermediateBudgetAndHonorsCancellation() {
        var inner = Caption("OVERWIDE CAPTION TEXT", OfficeTextAlignment.Center);
        var actual = new OfficeDrawing(20, 40).AddEffectDrawing(inner, OfficeTransform.Identity);
        actual.Fonts.Add(Family, Font());
        Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(actual,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 800 }));
        Assert.Throws<OperationCanceledException>(() => OfficeDrawingRasterRenderer.Render(actual,
            new OfficeDrawingRasterRenderOptions { CancellationToken = new System.Threading.CancellationToken(true) }));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void ExpandedLayerKeepsMaskRegistrationAndTheMasksOwnCanvasBoundary(int offset) {
        var inner = Caption("OVERWIDE CAPTION TEXT", OfficeTextAlignment.Center);
        var mask = new OfficeDrawing(40, 40);
        OfficeShape shape = OfficeShape.Rectangle(40, 40); shape.FillColor = OfficeColor.White; shape.StrokeWidth = 0;
        mask.AddShape(shape, 0, 0);
        OfficeTransform maskTransform = OfficeTransform.Translate(-20 + offset * 10, 0);
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(inner, OfficeTransform.Translate(140, 40),
            OfficeBlendMode.Normal, new OfficeDrawingSoftMask(mask, OfficeSoftMaskMode.Alpha, maskTransform));
        var expected = new OfficeDrawing(300, 160).AddClippedDrawing(inner, 120 + offset * 10, 40,
            OfficeClipPath.Rectangle(40, 40), 20 - offset * 10, 0);
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(actual);
        Assert.True(HasInkOutside(raster, 140, 160));
        EqualRaster(OfficeDrawingRasterRenderer.Render(expected), raster);
    }

    [Fact]
    public void UnwrappedMaskTextDoesNotExpandItsCoverageCanvas() {
        var inner = new OfficeDrawing(20, 40);
        OfficeShape shape = OfficeShape.Rectangle(20, 40); shape.FillColor = OfficeColor.Red; shape.StrokeWidth = 0;
        inner.AddShape(shape, 0, 0);
        var mask = Caption("OVERWIDE CAPTION TEXT", OfficeTextAlignment.Center);
        var actual = new OfficeDrawing(100, 80).AddEffectDrawing(inner, OfficeTransform.Translate(40, 20),
            OfficeBlendMode.Normal, new OfficeDrawingSoftMask(mask, OfficeSoftMaskMode.Alpha, OfficeTransform.Translate(20, 0)));
        actual.Fonts.Add(Family, Font());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(actual);
        Assert.All(Enumerable.Range(0, raster.Height), y => Assert.All(Enumerable.Range(0, raster.Width), x => Assert.Equal((byte)0, raster.GetPixel(x, y).A)));
    }

    [Fact]
    public void VerticallyInvisibleParagraphDoesNotEnlargeItsParentEffectSurface() {
        var inner = Caption("Fit", OfficeTextAlignment.Center);
        inner.AddEffectDrawing(Caption("OVERWIDE CAPTION TEXT", OfficeTextAlignment.Center), OfficeTransform.Translate(0, 100));
        var actual = new OfficeDrawing(20, 40).AddEffectDrawing(inner, OfficeTransform.Identity);
        actual.Fonts.Add(Family, Font());
        var expected = Caption("Fit", OfficeTextAlignment.Center); expected.Fonts.Add(Family, Font());
        EqualRaster(OfficeDrawingRasterRenderer.Render(expected), OfficeDrawingRasterRenderer.Render(actual,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 7000 }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OnlyActualPlacedLinesSizeAnAncestorEffect(bool mixedLines) {
        var child = new OfficeDrawing(20, 40).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(mixedLines ? "Fit\nOVERWIDE CAPTION TEXT" : "OVERWIDE CAPTION TEXT",
                10, OfficeColor.Black, fontFamily: Family) }, OfficeTextAlignment.Center)
        }, 0, 0, 20, 40, verticalAlignment: mixedLines ? OfficeTextVerticalAlignment.Top : OfficeTextVerticalAlignment.Bottom, wrapText: false);
        var inner = new OfficeDrawing(20, 40).AddEffectDrawing(child, OfficeTransform.Translate(0, 30));
        var actual = new OfficeDrawing(20, 40).AddEffectDrawing(inner, OfficeTransform.Identity);
        actual.Fonts.Add(Family, Font());
        var expected = new OfficeDrawing(20, 40);
        if (mixedLines) expected.AddEffectDrawing(Caption("Fit", OfficeTextAlignment.Center), OfficeTransform.Translate(0, 30));
        expected.Fonts.Add(Family, Font());
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(expected);
        Assert.Equal(mixedLines, Enumerable.Range(0, 40).Any(y => Enumerable.Range(0, 20).Any(x => reference.GetPixel(x, y).A > 0)));
        EqualRaster(reference, OfficeDrawingRasterRenderer.Render(actual));
        EqualRaster(reference, OfficeDrawingRasterRenderer.Render(actual,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 7000 }));
    }

    [Theory]
    [InlineData(OfficeTextDecorationStyle.Single)]
    [InlineData(OfficeTextDecorationStyle.Double)]
    [InlineData(OfficeTextDecorationStyle.Wavy)]
    public void VisibleBackgroundsAndDecorationsOutsideTheFrameAreRetained(OfficeTextDecorationStyle decoration) {
        var inner = new OfficeDrawing(20, 40).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun("OVERWIDE CAPTION TEXT", 10, OfficeColor.Black, italic: true, fontFamily: Family,
                    underlineStyle: decoration, strikethroughStyle: decoration, backgroundColor: OfficeColor.CornflowerBlue),
                new OfficeRichTextRun("       ", 10, OfficeColor.Black, fontFamily: Family, underlineStyle: decoration)
            }, OfficeTextAlignment.Center)
        }, 0, 0, 20, 40, verticalAlignment: OfficeTextVerticalAlignment.Bottom, wrapText: false);
        var actual = new OfficeDrawing(300, 160).AddEffectDrawing(inner, OfficeTransform.Translate(140, 40));
        var expected = new OfficeDrawing(300, 160).AddClippedDrawing(inner, 0, 40, OfficeClipPath.Rectangle(300, 40), 140, 0);
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(expected);
        Assert.True(HasInkOutside(reference, 140, 160));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(actual);
        int[] maximumDifference = new int[4]; int coverageDifference = 0;
        for (int y = 0; y < raster.Height; y++) for (int x = 0; x < raster.Width; x++) {
            OfficeColor expectedPixel = reference.GetPixel(x, y), actualPixel = raster.GetPixel(x, y);
            int[] difference = { Math.Abs(expectedPixel.R - actualPixel.R), Math.Abs(expectedPixel.G - actualPixel.G),
                Math.Abs(expectedPixel.B - actualPixel.B), Math.Abs(expectedPixel.A - actualPixel.A) };
            for (int channel = 0; channel < 4; channel++) maximumDifference[channel] = Math.Max(maximumDifference[channel], difference[channel]);
            if ((expectedPixel.A > 0) != (actualPixel.A > 0)) coverageDifference++;
        }
        // The isolated color layer rounds channels before the final composite.
        // Coverage is exact; double strokes can accumulate two blue-channel units.
        Assert.Equal(0, coverageDifference);
        Assert.InRange(maximumDifference[0], 0, 1); Assert.InRange(maximumDifference[1], 0, 1);
        Assert.InRange(maximumDifference[2], 0, decoration == OfficeTextDecorationStyle.Double ? 2 : 1);
        Assert.InRange(maximumDifference[3], 0, 1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TileLayersUseLateFontsAndRetainTheirIntrinsicViewport(bool overwide) {
        var inner = Caption(overwide ? "OVERWIDE CAPTION TEXT" : "Fit", OfficeTextAlignment.Center);
        var actual = new OfficeDrawing(300, 160).AddTilingPattern(inner,
            new OfficeImagePlacement(140, 40, 20, 40), 20, 40, repeatX: false, repeatY: false,
            transform: OfficeTransform.Translate(140, 40));
        var expected = new OfficeDrawing(300, 160).AddClippedDrawing(inner, 140, 40, OfficeClipPath.Rectangle(20, 40));
        actual.Fonts.Add(Family, Font()); expected.Fonts.Add(Family, Font());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(actual);
        EqualRaster(OfficeDrawingRasterRenderer.Render(expected), raster);
        Assert.False(HasInkOutside(raster, 140, 160));
        Assert.Empty(inner.Fonts.Faces);
        Assert.Empty(Assert.IsType<OfficeDrawingTilingPattern>(actual.Elements[0]).Tile.Fonts.Faces);
    }

    private static bool HasInkOutside(OfficeRasterImage raster, int left, int right) =>
        Enumerable.Range(0, raster.Height).Any(y => Enumerable.Range(0, raster.Width).Any(x => (x < left || x >= right) && raster.GetPixel(x, y).A > 0));
    private static void EqualRaster(OfficeRasterImage expected, OfficeRasterImage actual) {
        Assert.Equal((expected.Width, expected.Height), (actual.Width, actual.Height));
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++) Assert.Equal(expected.GetPixel(x, y), actual.GetPixel(x, y));
    }
    private sealed class CaptionGlyphProvider : IOfficeTextShapingProvider {
        private readonly int _advance;
        internal CaptionGlyphProvider(int advance) => _advance = advance;
        internal System.Collections.Generic.List<string?> Languages { get; } = new();
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Languages.Add(request.Language);
            return new OfficeTextShapingResult(request.Text.Select((value, index) =>
                new OfficeShapedGlyph(1, value.ToString(), index, advanceWidth: _advance)).ToArray());
        }
    }
}
