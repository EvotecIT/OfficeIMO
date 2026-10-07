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
    [InlineData(0, 20, 2, 19, 2)]
    [InlineData(90, 37, 20, 37, 19)]
    [InlineData(180, 19, 37, 20, 37)]
    [InlineData(270, 2, 19, 2, 20)]
    public void OfficeConicGradient_PreservesColorsOnBothSidesOfRotatedHardSeam(
        double angle, int afterX, int afterY, int beforeX, int beforeY) {
        var gradient = new OfficeConicGradient(0.5D, 0.5D, angle, new[] {
            new OfficeGradientStop(0D, OfficeColor.Red),
            new OfficeGradientStop(0.25D, OfficeColor.Red),
            new OfficeGradientStop(0.25D, OfficeColor.Blue),
            new OfficeGradientStop(1D, OfficeColor.Blue)
        });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(gradient.CreateDrawing(40D, 40D, 72));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(afterX, afterY));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(beforeX, beforeY));
        // The same hard boundary must hold near the center, where padded
        // triangle apexes previously painted into the neighboring color sector.
        for (int y = 18; y <= 21; y++) {
            for (int x = 18; x <= 21; x++) {
                double clockwise = Math.Atan2(x + 0.5D - 20D, 20D - (y + 0.5D)) * 180D / Math.PI;
                double relative = (clockwise - angle + 720D) % 360D;
                OfficeColor pixel = raster.GetPixel(x, y);
                Assert.True(relative < 90D ? pixel.R > pixel.B : pixel.B > pixel.R);
                Assert.True(pixel.A >= 200, "Antialiased wedges must retain center coverage.");
            }
        }
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

    [Fact]
    public void OfficeConicGradient_InterpolatesTransparentStopsInPremultipliedAlpha() {
        var gradient = new OfficeConicGradient(
            0.5D,
            0.5D,
            0D,
            new[] {
                new OfficeGradientStop(0D, OfficeColor.FromRgba(255, 0, 0, 0)),
                new OfficeGradientStop(1D, OfficeColor.FromRgba(0, 0, 255, 255))
            });

        OfficeColor halfway = gradient.Sample(0.5D);

        Assert.Equal(128, halfway.A);
        Assert.Equal(255, halfway.B);
        Assert.Equal(0, halfway.R);
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
