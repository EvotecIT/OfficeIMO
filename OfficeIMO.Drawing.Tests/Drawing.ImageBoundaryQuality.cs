using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingImageBoundaryQualityTests {
    private static readonly OfficeColor EdgeColor = OfficeColor.FromRgb(225, 30, 90);
    private static readonly OfficeTransform Rotation = OfficeTransform.Scale(1.5D, 1.5D)
        .Then(OfficeTransform.Translate(85D, 60D)).Then(OfficeTransform.RotateDegrees(33D, 157D, 108D));

    [Theory]
    [InlineData(0, 255)]
    [InlineData(1, 255)]
    [InlineData(2, 255)]
    [InlineData(0, 128)]
    [InlineData(1, 128)]
    [InlineData(2, 128)]
    public void RotatedBoundaryHasPartialCoverageWithoutChangingSourceColor(int route, byte alpha) {
        var color = OfficeColor.FromRgba(EdgeColor.R, EdgeColor.G, EdgeColor.B, alpha);
        OfficeRasterImage image = RenderRotated(new OfficeRasterImage(96, 64, color), route);
        Assert.Equal(color, image.GetPixel(157, 108));
        OfficeTransform inverse = Rotation.Invert();
        int partial = 0, outsideCentres = 0;
        double area = 0D;
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            OfficeColor pixel = image.GetPixel(x, y);
            area += pixel.A / (double)alpha;
            if (pixel.A == 0) continue;
            Assert.Equal(color.R, pixel.R);
            Assert.Equal(color.G, pixel.G);
            Assert.Equal(color.B, pixel.B);
            Assert.InRange(pixel.A, (byte)1, alpha);
            if (pixel.A < alpha) partial++;
            OfficePoint centre = inverse.TransformPoint(new OfficePoint(x + .5D, y + .5D));
            if (centre.X < 0D || centre.Y < 0D || centre.X >= 96D || centre.Y >= 64D) outsideCentres++;
        }
        Assert.InRange(area, 144D * 96D - 3D, 144D * 96D + 3D);
        Assert.True(partial > 400, "The transformed rectangle must cover boundary pixels partially.");
        Assert.True(outsideCentres > 100, "Pixels whose centres miss the image can still have covered area.");
    }

    [Fact]
    public void RotationAndAffineRoutesProduceTheSameBoundaryCoverage() {
        var source = new OfficeRasterImage(96, 64, EdgeColor);
        byte[] expected = RenderRotated(source, 0).GetPixels();
        Assert.Equal(expected, RenderRotated(source, 1).GetPixels());
        Assert.Equal(expected, RenderRotated(source, 2).GetPixels());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void TransparentTexelColorDoesNotFringeTheRotatedBoundary(int route) {
        var source = new OfficeRasterImage(96, 64, OfficeColor.Red);
        for (int y = 0; y < source.Height; y++) for (int x = 0; x < 16; x++) {
            source.SetPixel(x, y, OfficeColor.FromRgba(0, 0, 255, 0));
        }
        OfficeRasterImage image = RenderRotated(source, route);
        int partial = 0;
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            OfficeColor pixel = image.GetPixel(x, y);
            if (pixel.A == 0) continue;
            Assert.Equal(255, pixel.R);
            Assert.Equal(0, pixel.G);
            Assert.Equal(0, pixel.B);
            if (pixel.A < 255) partial++;
        }
        Assert.True(partial > 300);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void IntegerAlignedImageRetainsExactTexels(int route) {
        var source = new OfficeRasterImage(16, 8);
        for (int y = 0; y < source.Height; y++) for (int x = 0; x < source.Width; x++) {
            byte alpha = (byte)(((x + y) % 4) * 85);
            source.SetPixel(x, y, alpha == 0 ? OfficeColor.Transparent : OfficeColor.FromRgba((byte)(x * 13), (byte)(y * 29), 117, alpha));
        }
        var destination = new OfficeRasterImage(30, 20);
        var canvas = new OfficeRasterCanvas(destination);
        if (route == 0) canvas.DrawImage(source, 5D, 7D, 16D, 8D);
        else canvas.DrawAffineImage(source, OfficeTransform.Translate(5D, 7D), 1D,
            route == 1 ? OfficeBlendMode.Normal : OfficeBlendMode.Multiply);
        for (int y = 0; y < destination.Height; y++) for (int x = 0; x < destination.Width; x++) {
            Assert.Equal(x >= 5 && x < 21 && y >= 7 && y < 15 ? source.GetPixel(x - 5, y - 7) : OfficeColor.Transparent,
                destination.GetPixel(x, y));
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ImageSliverAtCanvasEdgeRetainsItsCoveredArea(int route) {
        var destination = new OfficeRasterImage(2, 1);
        var source = new OfficeRasterImage(1, 1, EdgeColor);
        var canvas = new OfficeRasterCanvas(destination);
        if (route == 0) canvas.DrawImage(source, -.4D, 0D, .8D, 1D);
        else canvas.DrawAffineImage(source, OfficeTransform.Scale(.8D, 1D).Then(OfficeTransform.Translate(-.4D, 0D)), 1D,
            route == 1 ? OfficeBlendMode.Normal : OfficeBlendMode.Multiply);
        Assert.Equal(OfficeColor.FromRgba(EdgeColor.R, EdgeColor.G, EdgeColor.B, 102), destination.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Transparent, destination.GetPixel(1, 0));
    }

    [Theory]
    [InlineData(OfficeBlendMode.Normal, 36, 21, 42)]
    [InlineData(OfficeBlendMode.Multiply, 10, 18, 32)]
    public void PartialCoverageCombinesSourceAlphaOpacityAndBackdrop(OfficeBlendMode blend, byte red, byte green, byte blue) {
        OfficeColor backdrop = OfficeColor.FromRgb(10, 20, 35);
        var destination = new OfficeRasterImage(2, 1, backdrop);
        var source = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(225, 30, 90, 128));
        new OfficeRasterCanvas(destination).DrawAffineImage(source,
            OfficeTransform.Scale(.8D, 1D).Then(OfficeTransform.Translate(-.4D, 0D)), .6D, blend);
        // The visible source covers 40% of the pixel: alpha rounds from 128 * .4 * .6 to 31.
        Assert.Equal(OfficeColor.FromRgb(red, green, blue), destination.GetPixel(0, 0));
        Assert.Equal(backdrop, destination.GetPixel(1, 0));
    }

    [Fact]
    public void ExplicitNearestNeighborRetainsPixelCentreBoundarySelection() {
        var destination = new OfficeRasterImage(2, 1);
        new OfficeRasterCanvas(destination).DrawImage(new OfficeRasterImage(1, 1, EdgeColor),
            new OfficeImageProjection(new OfficeImagePlacement(-.4D, 0D, .8D, 1D)), interpolate: false);
        Assert.Equal(OfficeColor.Transparent, destination.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Transparent, destination.GetPixel(1, 0));
    }

    [Theory]
    [InlineData(true, 102)]
    [InlineData(false, 0)]
    public void PublicDrawingRetainsThinImagesBetweenPixelCentres(bool interpolate, byte alpha) {
        byte[] encoded = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, EdgeColor));
        var drawing = new OfficeDrawing(2D, 1D).AddImageWithInterpolation(encoded, "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, .4D, 1D)), interpolate);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(alpha == 0 ? OfficeColor.Transparent : OfficeColor.FromRgba(EdgeColor.R, EdgeColor.G, EdgeColor.B, alpha), image.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Transparent, image.GetPixel(1, 0));
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public void ImageTouchingThePageWithoutCoveredAreaDoesNotAllocateAFilter(bool rightEdge, int route) {
        var destination = new OfficeRasterImage(2, 1);
        var canvas = new OfficeRasterCanvas(destination);
        canvas.ChargeIntermediateSurfacePixels(2L, 2L);
        var source = new OfficeRasterImage(16, 16, EdgeColor);
        double x = rightEdge ? 2D : -1D;
        if (route == 0) canvas.DrawImage(source, x, 0D, 1D, 1D);
        else canvas.DrawAffineImage(source, OfficeTransform.Scale(1D / 16D, 1D / 16D).Then(OfficeTransform.Translate(x, 0D)),
            1D, route == 1 ? OfficeBlendMode.Normal : OfficeBlendMode.Multiply);
        Assert.Equal(2L, canvas.TransformedTextBudget.IntermediatePixels);
        Assert.Equal(OfficeColor.Transparent, destination.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Transparent, destination.GetPixel(1, 0));
    }

    [Fact]
    public void ShearedBoundaryRetainsTheTransformedArea() {
        var image = new OfficeRasterImage(80, 80);
        new OfficeRasterCanvas(image).DrawAffineImage(new OfficeRasterImage(7, 5, EdgeColor),
            new OfficeTransform(1.5D, .3D, .6D, 1.1D, 30.2D, 40.4D));
        double area = 0D;
        int partial = 0;
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            byte alpha = image.GetPixel(x, y).A;
            area += alpha / 255D;
            if (alpha > 0 && alpha < 255) partial++;
        }
        Assert.InRange(area, 7D * 5D * (1.5D * 1.1D - .3D * .6D) - .15D, 7D * 5D * (1.5D * 1.1D - .3D * .6D) + .15D);
        Assert.True(partial > 20);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public void CanvasClipPreservesCoveredPixelsAndRejectsTheOutside(int route, bool transformedClip) {
        var source = new OfficeRasterImage(96, 64, EdgeColor);
        OfficeRasterImage expected = RenderRotated(source, route);
        var actual = new OfficeRasterImage(320, 190);
        var canvas = new OfficeRasterCanvas(actual);
        OfficeTransform clipTransform = OfficeTransform.RotateDegrees(-11D, 155D, 105D);
        OfficeTransform clipInverse = clipTransform.Invert();
        OfficePoint[] clipPoints = {
            clipTransform.TransformPoint(new OfficePoint(110D, 70D)),
            clipTransform.TransformPoint(new OfficePoint(200D, 70D)),
            clipTransform.TransformPoint(new OfficePoint(200D, 140D)),
            clipTransform.TransformPoint(new OfficePoint(110D, 140D))
        };
        using (transformedClip ? canvas.PushClipPolygon(clipPoints) : canvas.PushClipRectangle(110D, 70D, 90D, 70D)) {
            DrawRotated(canvas, source, route);
        }
        int partial = 0;
        for (int y = 0; y < actual.Height; y++) for (int x = 0; x < actual.Width; x++) {
            OfficeColor pixel = actual.GetPixel(x, y);
            OfficePoint centre = transformedClip ? clipInverse.TransformPoint(new OfficePoint(x + .5D, y + .5D)) : new OfficePoint(x + .5D, y + .5D);
            Assert.Equal(centre.X >= 110D && centre.X < 200D && centre.Y >= 70D && centre.Y < 140D
                ? expected.GetPixel(x, y) : OfficeColor.Transparent, pixel);
            if (pixel.A > 0 && pixel.A < 255) partial++;
        }
        Assert.True(partial > 30);
    }

    private static OfficeRasterImage RenderRotated(OfficeRasterImage source, int route) {
        var result = new OfficeRasterImage(320, 190);
        DrawRotated(new OfficeRasterCanvas(result), source, route);
        return result;
    }

    private static void DrawRotated(OfficeRasterCanvas canvas, OfficeRasterImage source, int route) {
        if (route == 0) canvas.DrawImage(source, 85D, 60D, 144D, 96D, 33D, 157D, 108D);
        else canvas.DrawAffineImage(source, Rotation, 1D, route == 1 ? OfficeBlendMode.Normal : OfficeBlendMode.Multiply);
    }
}
