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
