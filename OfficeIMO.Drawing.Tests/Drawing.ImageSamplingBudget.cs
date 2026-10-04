using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingImageSamplingBudgetTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void MinificationIncludesOutputAndFloatScratchInTheSharedPixelBudget(int route) {
        var source = new OfficeRasterImage(16, 16, OfficeColor.Red);
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(8, 8));
        canvas.ChargeIntermediateSurfacePixels(256, 831);
        void Draw() {
            if (route == 0) canvas.DrawImage(source, new OfficeImageProjection(new OfficeImagePlacement(0, 0, 8, 8)));
            else canvas.DrawAffineImage(source, OfficeTransform.Scale(.5D, .5D), 1D,
                route == 1 ? OfficeBlendMode.Normal : OfficeBlendMode.Multiply);
        }
        Assert.Throws<OfficeImageExportLimitException>(Draw);
        Assert.Equal(256, canvas.TransformedTextBudget.IntermediatePixels);

        var exact = new OfficeRasterCanvas(new OfficeRasterImage(8, 8));
        exact.ChargeIntermediateSurfacePixels(256, 832);
        if (route == 0) exact.DrawImage(source, new OfficeImageProjection(new OfficeImagePlacement(0, 0, 8, 8)));
        else exact.DrawAffineImage(source, OfficeTransform.Scale(.5D, .5D), 1D,
            route == 1 ? OfficeBlendMode.Normal : OfficeBlendMode.Multiply);
        // 256 decoded pixels + 64 destination pixels + 512 floats of four bytes each.
        Assert.Equal(832, exact.TransformedTextBudget.IntermediatePixels);
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
