using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingImageSamplingQualityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MinificationUsesTheSourcePixelAreaInCanvasAndDrawing(bool scene) {
        OfficeRasterImage source = Checker(300, 300);
        var expected = OfficeRasterResampler.Resize(source, 100, 100, OfficeRasterResamplingMode.Area);
        OfficeRasterImage actual;
        if (scene) {
            var drawing = new OfficeDrawing(100, 100);
            drawing.AddImage(OfficeRasterImageEncoder.Encode(source, OfficeImageExportFormat.Png), "image/png",
                new OfficeImageProjection(new OfficeImagePlacement(0, 0, 100, 100)));
            actual = OfficeDrawingRasterRenderer.Render(drawing);
        } else {
            actual = new OfficeRasterImage(100, 100);
            new OfficeRasterCanvas(actual).DrawImage(source, 0, 0, 100, 100);
        }
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }

    [Theory]
    [InlineData(OfficeBlendMode.Normal)]
    [InlineData(OfficeBlendMode.Multiply)]
    public void AffineCompositingUsesTheSameMinificationFilter(OfficeBlendMode blend) {
        OfficeRasterImage source = Checker(300, 300);
        var expected = OfficeRasterResampler.Resize(source, 100, 100, OfficeRasterResamplingMode.Area);
        var actual = new OfficeRasterImage(100, 100, OfficeColor.White);
        new OfficeRasterCanvas(actual).DrawAffineImage(source, OfficeTransform.Scale(1D / 3D, 1D / 3D), 1D, blend);
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }

    [Fact]
    public void CroppedMinificationKeepsItsPlacementAndExcludesTheOutsideColors() {
        var source = new OfficeRasterImage(600, 300, OfficeColor.Red);
        for (int y = 0; y < 300; y++) for (int x = 150; x < 450; x++) source.SetPixel(x, y, ((x + y) & 1) == 0 ? OfficeColor.Black : OfficeColor.White);
        var actual = new OfficeRasterImage(120, 120);
        new OfficeRasterCanvas(actual).DrawImage(source, 10, 10, 100, 100, .25, 0, .5, 1, 0, 60, 60, true, false);
        Assert.Equal(0, actual.GetPixel(9, 50).A);
        Assert.Equal(0, actual.GetPixel(110, 50).A);
        for (int y = 10; y < 110; y++) for (int x = 10; x < 110; x++) {
            OfficeColor pixel = actual.GetPixel(x, y);
            Assert.Equal(pixel.R, pixel.G);
            Assert.Equal(pixel.G, pixel.B);
            Assert.InRange(pixel.R, 113, 142);
        }
    }

    [Fact]
    public void MinificationFiltersPremultipliedAlphaWithoutTransparentColorHalos() {
        var source = new OfficeRasterImage(200, 200);
        for (int y = 0; y < 200; y++) for (int x = 0; x < 200; x++) source.SetPixel(x, y,
            ((x + y) & 1) == 0 ? OfficeColor.Red : OfficeColor.FromRgba(0, 0, 255, 0));
        var actual = new OfficeRasterImage(100, 100);
        new OfficeRasterCanvas(actual).DrawImage(source, 0, 0, 100, 100);
        Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 128), actual.GetPixel(50, 50));
    }

    [Fact]
    public void ExplicitNearestNeighborRetainsPixelArtDuringMinification() {
        var actual = new OfficeRasterImage(100, 100);
        new OfficeRasterCanvas(actual).DrawImage(Checker(300, 300), new OfficeImageProjection(new OfficeImagePlacement(0, 0, 100, 100)), interpolate: false);
        Assert.Equal(0, actual.GetPixel(0, 0).R);
        Assert.Equal(255, actual.GetPixel(1, 0).R);
    }

    private static OfficeRasterImage Checker(int width, int height) {
        var result = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) result.SetPixel(x, y, ((x + y) & 1) == 0 ? OfficeColor.Black : OfficeColor.White);
        return result;
    }
}
