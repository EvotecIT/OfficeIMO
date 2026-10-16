using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class PdfRasterConventionTests {
    [Fact]
    public void DuplicatePointsDoNotHideAReversingStroke() {
        var image = new OfficeRasterImage(100, 100, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image) { PdfRasterConventions = true };
        var points = new[] { new OfficePoint(20, 50), new OfficePoint(50, 50), new OfficePoint(50, 50), new OfficePoint(20, 50) };
        canvas.StrokeContours(new[] { new OfficeFlattenedPathContour(points, false) }, 12,
            OfficeStrokeLineCap.Butt, OfficeStrokeLineJoin.Miter, 10, null, 0, (_, _) => OfficeColor.Black);
        Assert.Equal(OfficeColor.White, image.GetPixel(55, 50));
        Assert.Equal(OfficeColor.Black, image.GetPixel(30, 50));
    }

    [Theory]
    [InlineData(792D, 1650)]
    [InlineData(1e-10D, 1)]
    public void RegionBoundsUseTheSamePixelGridAsTheFullRender(double height, int expectedHeight) {
        var drawing = new OfficeDrawing(612, height);
        var options = new OfficeDrawingRasterRenderOptions { Scale = 150D / 72D };
        OfficeRasterImage full = OfficeDrawingRasterRenderer.Render(drawing, options);
        Assert.Equal(expectedHeight, full.Height);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            OfficeDrawingRasterRenderer.RenderRegion(drawing, options, 0, full.Height, 1, 1));
        Assert.Equal(1, OfficeDrawingRasterRenderer.RenderRegion(drawing, options, 0, full.Height - 1, 1, 1).Height);
    }

    [Fact]
    public void IdenticalSoftMaskAndContentKeepTheirSnappedCoverage() {
        var source = new OfficeDrawing(10, 10);
        OfficeShape rectangle = OfficeShape.Rectangle(5.2D, 5.2D);
        rectangle.FillColor = OfficeColor.Red;
        rectangle.StrokeWidth = 0;
        source.AddShape(rectangle, 1.2D, 1.2D);
        var drawing = new OfficeDrawing(10, 10);
        drawing.AddEffectDrawing(source, OfficeTransform.Identity, OfficeBlendMode.Normal,
            new OfficeDrawingSoftMask(source, OfficeSoftMaskMode.Alpha));
        var options = new OfficeDrawingRasterRenderOptions { PdfRasterConventions = true };
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(source, options);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(drawing, options);
        for (int y = 0; y < expected.Height; y++)
            for (int x = 0; x < expected.Width; x++)
                Assert.Equal(expected.GetPixel(x, y), actual.GetPixel(x, y));
    }
}
