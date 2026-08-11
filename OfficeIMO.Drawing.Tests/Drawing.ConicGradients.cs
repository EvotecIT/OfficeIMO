using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingConicGradientTests {
    [Fact]
    public void OfficeConicGradient_ExpandsToAClippedBackendNeutralVectorDrawing() {
        var gradient = new OfficeConicGradient(
            0.5D,
            0.5D,
            0D,
            new[] {
                new OfficeGradientStop(0D, OfficeColor.Red),
                new OfficeGradientStop(0.25D, OfficeColor.Red),
                new OfficeGradientStop(0.25D, OfficeColor.Blue),
                new OfficeGradientStop(1D, OfficeColor.Blue)
            });

        OfficeDrawing drawing = gradient.CreateDrawing(40D, 40D, qualitySegments: 72);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Single(drawing.Elements);
        Assert.True(raster.GetPixel(20, 2).R > raster.GetPixel(20, 2).B);
        Assert.True(raster.GetPixel(37, 20).B > raster.GetPixel(37, 20).R);
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.True(Count(svg, "<path") >= 72);
        Assert.Equal(gradient.Stops, gradient.Clone().Stops);
    }

    [Theory]
    [InlineData(11)]
    [InlineData(4097)]
    public void OfficeConicGradient_BoundsVectorExpansion(int segments) {
        var gradient = new OfficeConicGradient(
            0.5D,
            0.5D,
            0D,
            new[] { new OfficeGradientStop(0D, OfficeColor.Red), new OfficeGradientStop(1D, OfficeColor.Blue) });
        Assert.Throws<ArgumentOutOfRangeException>(() => gradient.CreateDrawing(20D, 20D, segments));
    }

    [Fact]
    public void OfficeConicGradient_CoversTheBoxWhenTheAuthoredCenterIsOutsideIt() {
        var gradient = new OfficeConicGradient(
            5D,
            0.5D,
            0D,
            new[] { new OfficeGradientStop(0D, OfficeColor.Red), new OfficeGradientStop(1D, OfficeColor.Blue) });

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(gradient.CreateDrawing(40D, 20D, qualitySegments: 72));

        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(0, 0));
        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(39, 19));
    }

    [Fact]
    public void OfficeConicGradient_MinimumSegmentCountCoversCornersBetweenRays() {
        var gradient = new OfficeConicGradient(
            0.5D,
            0.5D,
            0D,
            new[] { new OfficeGradientStop(0D, OfficeColor.Red), new OfficeGradientStop(1D, OfficeColor.Blue) });

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(gradient.CreateDrawing(40D, 40D, qualitySegments: 12));

        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(0, 0));
        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(39, 0));
        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(0, 39));
        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(39, 39));
    }

    private static int Count(string value, string token) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += token.Length;
        }
        return count;
    }
}
