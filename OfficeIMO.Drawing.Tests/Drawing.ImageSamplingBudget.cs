using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingImageSamplingBudgetTests {
    [Fact]
    public void CachedPrefilterStillConsumesTheNewRendersWorkingBudget() {
        var source = new OfficeRasterImage(16, 16, OfficeColor.Red);
        var cache = new OfficeRasterImageCache(4096);
        var projection = new OfficeImageProjection(new OfficeImagePlacement(0, 0, 8, 8));
        var first = new OfficeRasterCanvas(new OfficeRasterImage(8, 8)) { ImageCache = cache };
        first.DrawImage(source, projection);
        Assert.Equal(256, cache.UsedBytes);

        var limited = new OfficeRasterCanvas(new OfficeRasterImage(8, 8)) { ImageCache = cache };
        limited.ChargeIntermediateSurfacePixels(256, 319);
        Assert.Throws<OfficeImageExportLimitException>(() => limited.DrawImage(source, projection));

        var output = new OfficeRasterImage(8, 8);
        var exact = new OfficeRasterCanvas(output) { ImageCache = cache };
        exact.ChargeIntermediateSurfacePixels(256, 320);
        exact.DrawImage(source, projection);
        Assert.Equal(320, exact.TransformedTextBudget.IntermediatePixels);
        Assert.Equal(OfficeColor.Red, output.GetPixel(4, 4));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void MinificationIncludesOutputAndFloatScratchInTheSharedPixelBudget(int route) {
        var source = new OfficeRasterImage(16, 16, OfficeColor.Red);
        OfficeRasterCanvas Draw(long maximumPixels) {
            var output = new OfficeRasterImage(8, 8);
            var canvas = new OfficeRasterCanvas(output);
            canvas.ChargeIntermediateSurfacePixels(256, maximumPixels);
            if (route == 0) canvas.DrawImage(source, new OfficeImageProjection(new OfficeImagePlacement(0, 0, 8, 8)));
            else canvas.DrawAffineImage(source, OfficeTransform.Scale(.5D, .5D), 1D,
                route == 1 ? OfficeBlendMode.Normal : OfficeBlendMode.Multiply);
            Assert.Equal(OfficeColor.Red, output.GetPixel(4, 4));
            return canvas;
        }

        // 256 decoded pixels + 64 destination pixels + 512 floats of four bytes each.
        Assert.Equal(832, Draw(832).TransformedTextBudget.IntermediatePixels);
        // Without room for the float scratch, whole 2x2 blocks are averaged instead:
        // 64 destination pixels + one row of eight four-channel double accumulators.
        Assert.Equal(384, Draw(831).TransformedTextBudget.IntermediatePixels);
        Assert.Equal(384, Draw(384).TransformedTextBudget.IntermediatePixels);
        Assert.Throws<OfficeImageExportLimitException>(() => Draw(383));
    }

    [Fact]
    public void EncodedImageMinificationEnforcesTheCallerLimitButNearestNeighborNeedsNoPrefilter() {
        byte[] encoded = OfficePngWriter.Encode(new OfficeRasterImage(16, 16, OfficeColor.Red));
        var projection = new OfficeImageProjection(new OfficeImagePlacement(0, 0, 8, 8));
        var drawing = new OfficeDrawing(8, 8).AddImage(encoded, "image/png", projection);
        Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 256 }));
        var nearest = new OfficeDrawing(8, 8).AddImageWithInterpolation(encoded, "image/png", projection, interpolate: false);
        OfficeRasterImage result = OfficeDrawingRasterRenderer.Render(nearest,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 256 });
        Assert.Equal(OfficeColor.Red, result.GetPixel(4, 4));
    }
}
