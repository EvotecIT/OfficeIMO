using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class PdfRasterConventionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedClipCoverageUsesTheGeometricIntersection(bool disjoint) {
        var image = new OfficeRasterImage(1, 1);
        var canvas = new OfficeRasterCanvas(image) { AntialiasedClipping = true };
        var left = new[] { new OfficePoint(0, 0), new OfficePoint(.5, 0), new OfficePoint(.5, 1), new OfficePoint(0, 1) };
        using (canvas.PushClipPolygon(left))
        using (canvas.PushClipRectangle(disjoint ? .5 : 0, 0, .5, 1))
            canvas.FillRectangle(0, 0, 1, 1, OfficeColor.Red);
        Assert.Equal(disjoint ? 0 : 128, image.GetPixel(0, 0).A);
    }

    [Fact]
    public void NestedNonrectangularClipsPreserveCoincidentCoverageAndHoles() {
        var triangle = new[] { new OfficePoint(0, 0), new OfficePoint(1, 0), new OfficePoint(0, 1) };
        var image = new OfficeRasterImage(1, 1);
        var canvas = new OfficeRasterCanvas(image) { AntialiasedClipping = true };
        using (canvas.PushClipPolygon(triangle))
        using (canvas.PushClipPolygonsNonZero(new[] { triangle }))
            canvas.FillRectangle(0, 0, 1, 1, OfficeColor.Red);
        Assert.Equal(128, image.GetPixel(0, 0).A);
        var holeImage = new OfficeRasterImage(1, 1);
        var holeCanvas = new OfficeRasterCanvas(holeImage) { AntialiasedClipping = true };
        using (holeCanvas.PushClipPolygon(triangle))
        using (holeCanvas.PushClipPolygonsEvenOdd(new[] { triangle, triangle }))
            holeCanvas.FillRectangle(0, 0, 1, 1, OfficeColor.Red);
        Assert.Equal(0, holeImage.GetPixel(0, 0).A);
    }

    [Fact]
    public void LargeFractionalExtentsStillRoundUpToTheNextPixel() {
        OfficeRasterScaleLimit limit = OfficeRasterScaleLimiter.Resolve(1000000000.4D, 1D, 1D, long.MaxValue);
        Assert.Equal(1000000001, limit.PixelWidth);
    }

    [Fact]
    public void BacktrackingClipDoesNotExposeItsBoundingRectangle() {
        var image = new OfficeRasterImage(3, 3);
        var canvas = new OfficeRasterCanvas(image) { PdfRasterConventions = true };
        using (canvas.PushClipPolygon(new[] { new OfficePoint(0, 0), new OfficePoint(2, 0),
            new OfficePoint(0, 0), new OfficePoint(0, 2) }))
            canvas.FillRectangle(0, 0, 3, 3, OfficeColor.Red);
        Assert.Equal(0, image.GetPixel(1, 1).A);
    }

    [Theory]
    [InlineData("a\nb", 2)]
    [InlineData("a\r\nb", 3)]
    [InlineData("a\u2028b", 2)]
    public void DisablingCjkBreaksPreservesMandatoryBoundaries(string text, int boundary) =>
        Assert.Equal(new[] { boundary }, OfficeTextLineBreaks.GetBreakPositions(text, false));

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
