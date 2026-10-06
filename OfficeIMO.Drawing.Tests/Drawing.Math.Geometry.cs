using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathGeometryTests {
    [Theory]
    [InlineData(72D, 20D, 14D)]
    [InlineData(96D, 80D / 3D, 56D / 3D)]
    public void ScopedGlyphMeasurementAndPaintUseTheSameDrawingUnits(double dpi, double em, double inkHeight) {
        var options = CreateOptions(ManagedTextShapingTestAssets.CreateFont('x'));
        options.Dpi = dpi;
        OfficeDrawing drawing = OfficeMathRenderer.Render(OfficeMath.Identifier("x"), options);
        OfficeDrawingText glyph = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());

        // The fixture is a 400-by-700-unit rectangle with a 500-unit advance in a 1000-unit em.
        Assert.Equal(em, glyph.Font.Size, 6);
        Assert.Equal(inkHeight, drawing.Height, 6);
        Assert.Equal(em / 2D, drawing.Width, 6);
        Assert.Equal(inkHeight, glyph.Font.Size + glyph.BaselineOffset, 6);
        Assert.Equal(em / 2D, glyph.TextAdvanceWidth!.Value, 6);
    }

    [Theory]
    [InlineData(OfficeFontStyle.Regular, 20D)]
    [InlineData(OfficeFontStyle.Bold | OfficeFontStyle.Italic, 20D)]
    [InlineData(OfficeFontStyle.Bold | OfficeFontStyle.Italic, 144D)]
    public void GlyphFrameIncludesPaintedOverhangsWithoutStretchingItsAdvance(OfficeFontStyle style, double size) {
        var options = CreateOptions(ManagedTextShapingTestAssets.CreateFont('x'));
        options.Font = new OfficeFontInfo("Scoped Math", size, style);
        OfficeDrawing drawing = OfficeMathRenderer.Render(OfficeMath.Identifier("x"), options);
        OfficeDrawingText glyph = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(size / 2D, glyph.TextAdvanceWidth!.Value, 6);
        Assert.Equal(size * .7D, drawing.Height, 6);
        Assert.True(drawing.Width >= size / 2D);

        // Compare visible ink with a generously framed positioned glyph. This exercises the
        // actual raster path, including synthetic styles, rather than only layout dimensions.
        var reference = new OfficeDrawing(size * 3D, size * 3D);
        reference.Fonts.AddRange(options.Fonts);
        reference.AddPositionedText("x", size, size, size, size, options.Font,
            OfficeColor.Black, textAdvanceWidth: size / 2D);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(drawing, scale: 4D);
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(reference, scale: 4D);
        int actualInk = CountInk(actual), expectedInk = CountInk(expected);
        Assert.InRange(actualInk, expectedInk - 8, expectedInk + 8);
    }

    [Fact]
    public void CompactNestedFractionsKeepDeclaredChildEmSizes() {
        var options = CreateOptions(ManagedTextShapingTestAssets.CreateFont('x', '2'));
        options.DisplayStyle = false;
        OfficeDrawing drawing = OfficeMathRenderer.Render(OfficeMath.Fraction(
            OfficeMath.Fraction(OfficeMath.Identifier("x"), OfficeMath.Number("2")),
            OfficeMath.Number("2")), options);
        OfficeDrawingText x = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), glyph => glyph.Text == "x");
        Assert.Equal(20D * 0.71D * 0.71D, x.Font.Size, 6);
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), glyph => Math.Abs(glyph.Font.Size - 14.2D) < 0.000001D);
        Assert.All(drawing.Elements.OfType<OfficeDrawingText>(), glyph => {
            Assert.InRange(glyph.X, 0D, drawing.Width - glyph.Width + 0.000001D);
            Assert.InRange(glyph.Y, 0D, drawing.Height - glyph.Height + 0.000001D);
        });
    }

    [Fact]
    public void ScopedColorLayersContributeToMathematicalInkBounds() {
        var options = CreateOptions(ManagedTextShapingTestAssets.CreateColorFont('x', baseGlyphHeight: 100));
        OfficeDrawing drawing = OfficeMathRenderer.Render(OfficeMath.Identifier("x"), options);
        // The monochrome base is only 100 units high; the two color layers are 700 units.
        Assert.Equal(14D, drawing.Height, 6);
        Assert.True(CountInk(OfficeDrawingRasterRenderer.Render(drawing, scale: 4D)) > 300);
    }

    [Fact]
    public void SmallScriptMeasurementDoesNotStretchGlyphsToTheGeneralTextFloor() {
        var options = CreateOptions(ManagedTextShapingTestAssets.CreateFont('x'));
        options.Font = options.Font.WithSize(0.5D);
        OfficeMathLayoutMetrics measured = OfficeMathRenderer.Measure(OfficeMath.Identifier("x"), options);
        Assert.Equal(0.25D, measured.Width, 6);
        Assert.Equal(0.35D, measured.Height, 6);
    }

    [Fact]
    public void SmallScriptPaintUsesItsDeclaredEmAndBaseline() {
        var options = CreateOptions(ManagedTextShapingTestAssets.CreateFont('x'));
        options.Font = options.Font.WithSize(0.5D);
        OfficeDrawing math = OfficeMathRenderer.Render(OfficeMath.Identifier("x"), options);
        var glyph = Assert.Single(math.Elements.OfType<OfficeDrawingText>()).Clone();
        var drawing = new OfficeDrawing(30D, 30D);
        drawing.Fonts.AddRange(options.Fonts);
        drawing.AddPositionedText(glyph.Text, 10D, 10D, glyph.Width, glyph.Height, glyph.Font, OfficeColor.Black,
            glyph.Alignment, lineHeight: null, textAdvanceWidth: glyph.TextAdvanceWidth,
            underlineStyle: OfficeTextDecorationStyle.None, strikethroughStyle: OfficeTextDecorationStyle.None,
            baseline: glyph.Baseline, baselineLevel: glyph.BaselineLevel, baselineScale: glyph.BaselineScale, baselineOffset: glyph.BaselineOffset);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        // A .5-unit em paints a .2 by .35 rectangle. Supersampled coverage must
        // remain bounded by that subpixel ink, rather than a one-unit replacement.
        int coverage = 0;
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++) coverage += raster.GetPixel(x, y).A;
        Assert.InRange(coverage, 1, 35);
    }

    [Theory]
    [InlineData(20D)]
    [InlineData(0.5D)]
    public void MixedScopedFacesShareThePaintedBaselineAndNaturalAdvance(double size) {
        var options = CreateOptions(ManagedTextShapingTestAssets.CreateFont('x'));
        options.Font = new OfficeFontInfo("Scoped Math, Scoped Digits", size);
        options.Fonts.Add("Scoped Digits", ManagedTextShapingTestAssets.CreateFont('2'));
        OfficeMathLayoutMetrics measured = OfficeMathRenderer.Measure(OfficeMath.Text("x2"), options);
        Assert.Equal(size, measured.Width, 6);
        Assert.Equal(size * 0.7D, measured.Height, 6);
    }

    private static OfficeMathRenderOptions CreateOptions(byte[] data) {
        var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Scoped Math", 20D), Padding = 0D };
        options.Fonts.Add("Scoped Math", data);
        return options;
    }

    private static int CountInk(OfficeRasterImage raster) {
        int count = 0;
        for (int y = 0; y < raster.Height; y++) {
            for (int x = 0; x < raster.Width; x++) {
                if (raster.GetPixel(x, y).A > 127) count++;
            }
        }
        return count;
    }
}
