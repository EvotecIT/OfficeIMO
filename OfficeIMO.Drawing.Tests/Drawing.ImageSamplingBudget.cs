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
        // Without room for the float scratch, the source is averaged one row at a time instead:
        // 64 destination pixels + three rows of eight four-channel double accumulators.
        Assert.Equal(512, Draw(831).TransformedTextBudget.IntermediatePixels);
        Assert.Equal(512, Draw(512).TransformedTextBudget.IntermediatePixels);
        Assert.Throws<OfficeImageExportLimitException>(() => Draw(511));
    }

    [Theory]
    [InlineData(1001, 1001, 2, 2)]
    [InlineData(1003, 701, 3, 2)]
    public void BudgetLimitedMinificationKeepsPartialEdgeCoverage(int sourceWidth, int sourceHeight, int width, int height) {
        // White with a one-pixel black right column and bottom row. Neither size divides evenly
        // into the output, so the edge output pixels cover a fractional share of the source.
        var source = new OfficeRasterImage(sourceWidth, sourceHeight, OfficeColor.White);
        for (int y = 0; y < sourceHeight; y++) source.SetPixel(sourceWidth - 1, y, OfficeColor.Black);
        for (int x = 0; x < sourceWidth; x++) source.SetPixel(x, sourceHeight - 1, OfficeColor.Black);
        long sourcePixels = (long)sourceWidth * sourceHeight;
        OfficeRasterImage Draw(long workingPixels, out long charged) {
            var output = new OfficeRasterImage(width, height);
            var canvas = new OfficeRasterCanvas(output);
            canvas.ChargeIntermediateSurfacePixels(sourcePixels, sourcePixels + workingPixels);
            canvas.DrawImage(source, new OfficeImageProjection(new OfficeImagePlacement(0, 0, width, height)));
            charged = canvas.TransformedTextBudget.IntermediatePixels - sourcePixels;
            return output;
        }

        OfficeRasterImage single = Draw(1_000_000L, out long singleCharge);
        OfficeRasterImage streamed = Draw(width * height + width * 24L, out long streamedCharge);

        Assert.True(streamedCharge < singleCharge);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                OfficeColor expected = single.GetPixel(x, y), actual = streamed.GetPixel(x, y);
                Assert.InRange(actual.R - expected.R, -1, 1);
                Assert.InRange(actual.A - expected.A, -1, 1);
            }
        }
        // The black edge is a sliver of the corner pixel, not a third of it.
        Assert.True(streamed.GetPixel(width - 1, height - 1).R > 245);
    }

    [Fact]
    public void BudgetLimitedMinificationCountsTheCanvasAndSourceInTheManagedCeiling() {
        // 5000x8000 source (160 MB) reduced to 1000x1600 beside a retained canvas.
        Assert.True(OfficeRasterResampler.TryGetStreamingAreaWorkingSetBytes(
            5000, 8000, 1000, 1600, retainedManagedBytes: 4L * 1000 * 1600, out long small));
        Assert.True(small <= OfficeRasterGuards.MaximumDecodedBytes);
        // A 10000x5000 canvas (200 MB) leaves no room for the same source.
        Assert.False(OfficeRasterResampler.TryGetStreamingAreaWorkingSetBytes(
            5000, 8000, 1000, 1600, retainedManagedBytes: 4L * 10000 * 5000, out long large));
        Assert.True(large > OfficeRasterGuards.MaximumDecodedBytes);
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
