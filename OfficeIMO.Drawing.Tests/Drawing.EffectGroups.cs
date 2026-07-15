using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDrawingEffectGroup_CompositesOpacityOnceAfterTransform() {
        var inner = new OfficeDrawing(20D, 20D);
        OfficeShape first = OfficeShape.Rectangle(10D, 10D);
        first.FillColor = OfficeColor.Red;
        first.StrokeWidth = 0D;
        OfficeShape second = first.Clone();
        inner.AddShape(first, 0D, 0D);
        inner.AddShape(second, 0D, 0D);

        var drawing = new OfficeDrawing(40D, 30D);
        drawing.AddEffectDrawing(inner, OfficeTransform.Translate(10D, 5D), 0.5D);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        OfficeColor painted = raster.GetPixel(12, 7);

        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(2, 2));
        Assert.Equal((byte)255, painted.R);
        Assert.Equal((byte)0, painted.G);
        Assert.Equal((byte)0, painted.B);
        Assert.InRange(painted.A, (byte)127, (byte)128);
        Assert.Contains("opacity=\"0.5\"", svg, StringComparison.Ordinal);
        Assert.Contains("matrix(1 0 0 1 10 5)", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingEffectGroup_AppliesArbitraryAffineScale() {
        var inner = new OfficeDrawing(10D, 10D);
        OfficeShape shape = OfficeShape.Rectangle(10D, 10D);
        shape.FillColor = OfficeColor.Blue;
        shape.StrokeWidth = 0D;
        inner.AddShape(shape, 0D, 0D);

        var drawing = new OfficeDrawing(40D, 30D);
        OfficeTransform transform = OfficeTransform.Scale(2D, 1.5D).Then(OfficeTransform.Translate(5D, 4D));
        drawing.AddEffectDrawing(inner, transform);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.Equal(OfficeColor.Blue, raster.GetPixel(6, 5));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(23, 17));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(26, 17));
    }

    [Fact]
    public void OfficeDrawingEffectGroup_AppliesAffineRotation() {
        var inner = new OfficeDrawing(10D, 20D);
        OfficeShape shape = OfficeShape.Rectangle(10D, 20D);
        shape.FillColor = OfficeColor.Blue;
        shape.StrokeWidth = 0D;
        inner.AddShape(shape, 0D, 0D);

        var drawing = new OfficeDrawing(35D, 30D);
        OfficeTransform transform = OfficeTransform.RotateDegrees(90D).Then(OfficeTransform.Translate(25D, 5D));
        drawing.AddEffectDrawing(inner, transform);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(OfficeColor.Blue, raster.GetPixel(6, 6));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(23, 13));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(10, 20));
        Assert.Contains("matrix(0 1 -1 0 25 5)", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingEffectGroup_UsesManagedMultiplyBlending() {
        var source = new OfficeDrawing(8D, 8D);
        OfficeShape blue = OfficeShape.Rectangle(8D, 8D);
        blue.FillColor = OfficeColor.FromRgb(64, 128, 255);
        blue.StrokeWidth = 0D;
        source.AddShape(blue, 0D, 0D);

        var drawing = new OfficeDrawing(8D, 8D);
        OfficeShape orange = OfficeShape.Rectangle(8D, 8D);
        orange.FillColor = OfficeColor.FromRgb(240, 128, 32);
        orange.StrokeWidth = 0D;
        drawing.AddShape(orange, 0D, 0D);
        drawing.AddEffectDrawing(source, OfficeTransform.Identity, OfficeBlendMode.Multiply);

        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(4, 4);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.InRange(pixel.R, (byte)59, (byte)61);
        Assert.InRange(pixel.G, (byte)63, (byte)65);
        Assert.InRange(pixel.B, (byte)31, (byte)33);
        Assert.Equal((byte)255, pixel.A);
        Assert.Contains("mix-blend-mode:multiply", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingEffectGroup_AppliesReusableLuminositySoftMask() {
        var source = new OfficeDrawing(10D, 4D);
        OfficeShape red = OfficeShape.Rectangle(10D, 4D);
        red.FillColor = OfficeColor.Red;
        red.StrokeWidth = 0D;
        source.AddShape(red, 0D, 0D);

        var maskDrawing = new OfficeDrawing(10D, 4D);
        OfficeShape whiteHalf = OfficeShape.Rectangle(5D, 4D);
        whiteHalf.FillColor = OfficeColor.White;
        whiteHalf.StrokeWidth = 0D;
        maskDrawing.AddShape(whiteHalf, 0D, 0D);
        var mask = new OfficeDrawingSoftMask(maskDrawing, OfficeSoftMaskMode.Luminosity);

        var drawing = new OfficeDrawing(10D, 4D);
        drawing.AddEffectDrawing(source, OfficeTransform.Identity, OfficeBlendMode.Normal, mask);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        OfficeDrawingEffectGroup effect = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(2, 2));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(8, 2));
        Assert.Equal(OfficeSoftMaskMode.Luminosity, effect.SoftMask!.Mode);
        Assert.Contains("<mask id=\"officeimo-mask-", svg, StringComparison.Ordinal);
        Assert.Contains("mask-type:luminance", svg, StringComparison.Ordinal);
        Assert.Contains("mask=\"url(#officeimo-mask-", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingEffectGroup_CompositesPartialLuminosityMaskOverBackdrop() {
        var source = new OfficeDrawing(4D, 4D);
        OfficeShape red = OfficeShape.Rectangle(4D, 4D);
        red.FillColor = OfficeColor.Red;
        red.StrokeWidth = 0D;
        source.AddShape(red, 0D, 0D);

        var maskDrawing = new OfficeDrawing(4D, 4D);
        OfficeShape translucentBlack = OfficeShape.Rectangle(4D, 4D);
        translucentBlack.FillColor = OfficeColor.FromRgba(0, 0, 0, 128);
        translucentBlack.StrokeWidth = 0D;
        maskDrawing.AddShape(translucentBlack, 0D, 0D);
        var mask = new OfficeDrawingSoftMask(
            maskDrawing,
            OfficeSoftMaskMode.Luminosity,
            backdropColor: OfficeColor.White);

        var drawing = new OfficeDrawing(4D, 4D);
        drawing.AddEffectDrawing(source, OfficeTransform.Identity, OfficeBlendMode.Normal, mask);

        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(2, 2);

        Assert.Equal((byte)255, pixel.R);
        Assert.InRange(pixel.A, (byte)126, (byte)128);
    }
}
