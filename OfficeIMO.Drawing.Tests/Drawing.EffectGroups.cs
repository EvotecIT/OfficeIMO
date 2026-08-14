using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDrawingSoftMask_RejectsUndefinedPublicEnumValues() {
        var maskDrawing = new OfficeDrawing(1D, 1D);

        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDrawingSoftMask(
            maskDrawing,
            (OfficeSoftMaskMode)99));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDrawingSoftMask(
            maskDrawing,
            luminosityStandard: (OfficeSoftMaskLuminosityStandard)99));
    }

    [Fact]
    public void OfficeDrawingSvgExporter_VectorizesInexactNearestNeighborScale() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
        var inexact = new OfficeDrawing(10D, 10D);
        inexact.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 5D, 3D)),
            interpolate: false);
        var exact = new OfficeDrawing(10D, 10D);
        exact.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 6D, 4D)),
            interpolate: false);

        string inexactSvg = OfficeDrawingSvgExporter.ToSvg(inexact);
        string svg = OfficeDrawingSvgExporter.ToSvg(exact);

        Assert.Contains("shape-rendering=\"crispEdges\"", inexactSvg, StringComparison.Ordinal);
        Assert.Contains("scale(2.5 1.5)", inexactSvg, StringComparison.Ordinal);
        Assert.DoesNotContain("<image", inexactSvg, StringComparison.Ordinal);
        Assert.Contains("shape-rendering=\"crispEdges\"", svg, StringComparison.Ordinal);
    }

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

    [Theory]
    [InlineData(OfficeBlendMode.Normal)]
    [InlineData(OfficeBlendMode.Multiply)]
    public void OfficeDrawingEffectGroup_PreservesNearestNeighborSamplingThroughTransform(OfficeBlendMode blendMode) {
        OfficeRasterImage source = new OfficeRasterImage(2, 1, OfficeColor.Transparent);
        source.SetPixel(0, 0, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        byte[] png = OfficePngWriter.Encode(source);
        var inner = new OfficeDrawing(2D, 1D);
        inner.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: false);

        var drawing = new OfficeDrawing(4D, 1D);
        drawing.AddEffectDrawing(inner, OfficeTransform.Scale(2D, 1D), blendMode);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.Equal(OfficeColor.Black, raster.GetPixel(1, 0));
        Assert.Equal(OfficeColor.White, raster.GetPixel(2, 0));
    }

    [Theory]
    [InlineData(OfficeBlendMode.Normal)]
    [InlineData(OfficeBlendMode.Multiply)]
    public void OfficeDrawingEffectGroup_IgnoresTransparentNearestNeighborImagesForSampling(OfficeBlendMode blendMode) {
        OfficeRasterImage source = new OfficeRasterImage(2, 1, OfficeColor.Transparent);
        source.SetPixel(0, 0, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        byte[] png = OfficePngWriter.Encode(source);
        var inner = new OfficeDrawing(2D, 1D);
        inner.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: true);
        inner.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: false,
            opacity: 0D);

        var drawing = new OfficeDrawing(4D, 1D);
        drawing.AddEffectDrawing(inner, OfficeTransform.Scale(2D, 1D), blendMode);

        OfficeColor boundary = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(1, 0);

        Assert.InRange(boundary.R, (byte)1, (byte)254);
        Assert.Equal(boundary.R, boundary.G);
        Assert.Equal(boundary.R, boundary.B);
    }

    [Theory]
    [InlineData(OfficeBlendMode.Normal)]
    [InlineData(OfficeBlendMode.Multiply)]
    public void OfficeDrawingEffectGroup_IgnoresClippedNearestNeighborImagesForSampling(OfficeBlendMode blendMode) {
        OfficeRasterImage source = new OfficeRasterImage(2, 1, OfficeColor.Transparent);
        source.SetPixel(0, 0, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        byte[] png = OfficePngWriter.Encode(source);
        var clipped = new OfficeDrawing(2D, 1D);
        clipped.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: false);
        var inner = new OfficeDrawing(2D, 1D);
        inner.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: true);
        inner.AddClippedDrawing(clipped, 0D, 0D, OfficeClipPath.Rectangle(2D, 1D), 3D, 0D);

        var drawing = new OfficeDrawing(4D, 1D);
        drawing.AddEffectDrawing(inner, OfficeTransform.Scale(2D, 1D), blendMode);

        OfficeColor boundary = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(1, 0);

        Assert.InRange(boundary.R, (byte)1, (byte)254);
        Assert.Equal(boundary.R, boundary.G);
        Assert.Equal(boundary.R, boundary.B);
    }

    [Theory]
    [InlineData(OfficeBlendMode.Normal)]
    [InlineData(OfficeBlendMode.Multiply)]
    public void OfficeDrawingEffectGroup_IgnoresClippedNearestNeighborImagesInsideTransformedGroups(OfficeBlendMode blendMode) {
        OfficeRasterImage source = new OfficeRasterImage(2, 1, OfficeColor.Transparent);
        source.SetPixel(0, 0, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        byte[] png = OfficePngWriter.Encode(source);
        var clipped = new OfficeDrawing(2D, 1D);
        clipped.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: false);
        var inner = new OfficeDrawing(2D, 1D);
        inner.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: true);
        inner.AddClippedDrawing(
            clipped,
            0D,
            0D,
            OfficeClipPath.Rectangle(2D, 1D),
            3D,
            0D,
            new OfficeImageFrameTransform(180D, 1D, 0.5D));

        var drawing = new OfficeDrawing(4D, 1D);
        drawing.AddEffectDrawing(inner, OfficeTransform.Scale(2D, 1D), blendMode);

        OfficeColor boundary = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(1, 0);

        Assert.InRange(boundary.R, (byte)1, (byte)254);
        Assert.Equal(boundary.R, boundary.G);
        Assert.Equal(boundary.R, boundary.B);
    }

    [Theory]
    [InlineData(OfficeBlendMode.Normal)]
    [InlineData(OfficeBlendMode.Multiply)]
    public void OfficeDrawingEffectGroup_PreservesNearestNeighborSamplingFromSoftMask(OfficeBlendMode blendMode) {
        var source = new OfficeDrawing(2D, 1D);
        OfficeShape red = OfficeShape.Rectangle(2D, 1D);
        red.FillColor = OfficeColor.Red;
        red.StrokeWidth = 0D;
        source.AddShape(red, 0D, 0D);

        OfficeRasterImage maskPixels = new OfficeRasterImage(2, 1, OfficeColor.Transparent);
        maskPixels.SetPixel(0, 0, OfficeColor.Black);
        maskPixels.SetPixel(1, 0, OfficeColor.White);
        var maskDrawing = new OfficeDrawing(2D, 1D);
        maskDrawing.AddImage(
            OfficePngWriter.Encode(maskPixels),
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: false);
        var mask = new OfficeDrawingSoftMask(maskDrawing, OfficeSoftMaskMode.Luminosity);

        var drawing = new OfficeDrawing(4D, 1D);
        drawing.AddEffectDrawing(source, OfficeTransform.Scale(2D, 1D), blendMode, mask);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(1, 0));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(2, 0));
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
    public void OfficeDrawingEffectGroup_IsolatesEverySvgCompositionBoundary() {
        var painted = new OfficeDrawing(8D, 8D);
        OfficeShape red = OfficeShape.Rectangle(8D, 8D);
        red.FillColor = OfficeColor.Red;
        red.StrokeWidth = 0D;
        painted.AddShape(red, 0D, 0D);

        var nested = new OfficeDrawing(8D, 8D);
        nested.AddEffectDrawing(painted, OfficeTransform.Identity, OfficeBlendMode.Multiply);
        var drawing = new OfficeDrawing(8D, 8D);
        drawing.AddEffectDrawing(nested, OfficeTransform.Identity);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(2, svg.Split(new[] { "isolation:isolate" }, StringSplitOptions.None).Length - 1);
        Assert.Contains("isolation:isolate;mix-blend-mode:multiply", svg, StringComparison.Ordinal);
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

    [Fact]
    public void OfficeDrawingEffectGroup_PreservesPdfDeviceRgbLuminosityAcrossRasterAndSvg() {
        var source = new OfficeDrawing(4D, 4D);
        OfficeShape sourceShape = OfficeShape.Rectangle(4D, 4D);
        sourceShape.FillColor = OfficeColor.Blue;
        sourceShape.StrokeWidth = 0D;
        source.AddShape(sourceShape, 0D, 0D);

        var maskDrawing = new OfficeDrawing(4D, 4D);
        OfficeShape redMask = OfficeShape.Rectangle(4D, 4D);
        redMask.FillColor = OfficeColor.Red;
        redMask.StrokeWidth = 0D;
        maskDrawing.AddShape(redMask, 0D, 0D);
        var mask = new OfficeDrawingSoftMask(
            maskDrawing,
            OfficeSoftMaskMode.Luminosity,
            luminosityStandard: OfficeSoftMaskLuminosityStandard.PdfDeviceRgb);

        var drawing = new OfficeDrawing(4D, 4D);
        drawing.AddEffectDrawing(source, OfficeTransform.Identity, OfficeBlendMode.Normal, mask);

        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(2, 2);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.InRange(pixel.A, (byte)76, (byte)77);
        Assert.Contains("0.3 0.59 0.11", svg, StringComparison.Ordinal);
    }
}
