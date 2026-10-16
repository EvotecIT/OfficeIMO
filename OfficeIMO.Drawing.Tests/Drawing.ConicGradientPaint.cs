using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingConicGradientPaintTests {
    [Theory]
    [InlineData(12, 255, 0.5, 0.5)]
    [InlineData(360, 255, 0.5, 0.5)]
    [InlineData(360, 128, 0.5, 0.5)]
    [InlineData(4096, 128, -0.3, 2.0)]
    public void UniformConicPaintCoversTheBoxExactlyOnce(
        int segments, byte alpha, double centerX, double centerY) {
        OfficeColor color = OfficeColor.FromRgba(255, 0, 0, alpha);
        var gradient = new OfficeConicGradient(centerX, centerY, 23D, new[] {
            new OfficeGradientStop(0D, color),
            new OfficeGradientStop(0.25D, color),
            new OfficeGradientStop(1D, color)
        });

        OfficeDrawing drawing = gradient.CreateDrawing(100D, 80D, segments);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        for (int y = 0; y < raster.Height; y++) {
            for (int x = 0; x < raster.Width; x++) {
                Assert.True(raster.GetPixel(x, y) == color,
                    $"Uniform conic paint changed color or alpha at ({x},{y}): {raster.GetPixel(x, y)}.");
            }
        }
        Assert.DoesNotContain("<image", OfficeDrawingSvgExporter.ToSvg(drawing), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 255, 255)]
    [InlineData(17, 128, 128)]
    [InlineData(17, 64, 192)]
    public void FlatConicSectorsRetainInteriorColorAndAlpha(
        double angle, byte redAlpha, byte blueAlpha) {
        OfficeColor red = OfficeColor.FromRgba(255, 0, 0, redAlpha);
        OfficeColor blue = OfficeColor.FromRgba(0, 0, 255, blueAlpha);
        var gradient = new OfficeConicGradient(0.5D, 0.5D, angle, new[] {
            new OfficeGradientStop(0D, red),
            new OfficeGradientStop(0.25D, red),
            new OfficeGradientStop(0.25D, blue),
            new OfficeGradientStop(1D, blue)
        });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(gradient.CreateDrawing(100D, 80D));

        for (int y = 1; y < raster.Height - 1; y++) {
            for (int x = 1; x < raster.Width - 1; x++) {
                double dx = x + 0.5D - 50D, dy = 40D - y - 0.5D;
                double clockwise = Math.Atan2(dx, dy) * 180D / Math.PI;
                double relative = (clockwise - angle + 720D) % 360D;
                double radius = Math.Sqrt(dx * dx + dy * dy);
                // True color boundaries retain antialiasing. Internal quality rays
                // must not change the paint, including within translucent sectors.
                double boundaryDistance = radius * Math.Min(
                    Math.Abs(Math.Sin(relative * Math.PI / 180D)),
                    Math.Abs(Math.Sin((relative - 90D) * Math.PI / 180D)));
                if (boundaryDistance < 2D) continue;
                OfficeColor expected = relative < 90D ? red : blue;
                Assert.True(raster.GetPixel(x, y) == expected,
                    $"Flat conic sector changed color or alpha at ({x},{y}): {raster.GetPixel(x, y)}.");
            }
        }
    }

    [Fact]
    public void SmoothConicPaintAppliesACommonAlphaOnce() {
        const byte alpha = 128;
        var gradient = new OfficeConicGradient(0.5D, 0.5D, 13D, new[] {
            new OfficeGradientStop(0D, OfficeColor.FromRgba(255, 0, 0, alpha)),
            new OfficeGradientStop(0.5D, OfficeColor.FromRgba(0, 0, 255, alpha)),
            new OfficeGradientStop(1D, OfficeColor.FromRgba(255, 0, 0, alpha))
        });
        OfficeDrawing drawing = gradient.CreateDrawing(100D, 80D);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        for (int y = 0; y < raster.Height; y++) {
            for (int x = 0; x < raster.Width; x++) {
                // The complete opaque color field receives its shared alpha
                // once, including the center, quality rays, and box edges.
                Assert.Equal(alpha, raster.GetPixel(x, y).A);
            }
        }
        Assert.DoesNotContain("<image", OfficeDrawingSvgExporter.ToSvg(drawing), StringComparison.Ordinal);
    }

    [Fact]
    public void VaryingConicAlphaIsSampledOnceAndKeepsPremultipliedColor() {
        var gradient = new OfficeConicGradient(0.5D, 0.5D, 0D, new[] {
            new OfficeGradientStop(0D, OfficeColor.FromRgba(255, 0, 0, 0)),
            new OfficeGradientStop(0.5D, OfficeColor.Blue),
            new OfficeGradientStop(1D, OfficeColor.FromRgba(255, 0, 0, 0))
        });
        OfficeDrawing drawing = gradient.CreateDrawing(100D, 80D);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        var witnesses = new[] { (80, 10), (90, 55), (20, 65), (5, 20), (50, 10), (10, 40), (75, 65), (25, 10) };
        foreach ((int x, int y) in witnesses) {
            double turn = (Math.Atan2(x + 0.5D - 50D, 40D - y - 0.5D) / (Math.PI * 2D) + 1D) % 1D;
            int expectedAlpha = (int)Math.Round((turn <= 0.5D ? turn * 2D : (1D - turn) * 2D) * 255D,
                MidpointRounding.AwayFromZero);
            OfficeColor pixel = raster.GetPixel(x, y);
            Assert.Equal((byte)0, pixel.R);
            Assert.Equal((byte)0, pixel.G);
            Assert.Equal((byte)255, pixel.B);
            Assert.InRange((int)pixel.A, Math.Max(0, expectedAlpha - 3), Math.Min(255, expectedAlpha + 3));
        }
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("<mask", svg, StringComparison.Ordinal);
        Assert.Contains("mask-type:luminance", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("<image", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void FlatConicRegionCrossingTheOriginalTurnKeepsItsColorAndAlpha() {
        OfficeColor red = OfficeColor.FromRgba(255, 0, 0, 64);
        OfficeColor blue = OfficeColor.FromRgba(0, 0, 255, 192);
        var gradient = new OfficeConicGradient(0.5D, 0.5D, 0D, new[] {
            new OfficeGradientStop(0D, red),
            new OfficeGradientStop(0.25D, red),
            new OfficeGradientStop(0.25D, blue),
            new OfficeGradientStop(0.5D, blue),
            new OfficeGradientStop(0.5D, red),
            new OfficeGradientStop(1D, red)
        });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(gradient.CreateDrawing(100D, 80D));
        // The upward turn boundary lies inside one uninterrupted red sector.
        // A private tessellation cut must not create a color or opacity seam.
        for (int y = 1; y <= 30; y++) {
            for (int x = 45; x <= 54; x++) Assert.Equal(red, raster.GetPixel(x, y));
        }
        Assert.Equal(blue, raster.GetPixel(80, 60));
    }

    [Theory]
    [InlineData(0.0, 0.0)]
    [InlineData(0.0, 0.5)]
    [InlineData(1.0, 1.0)]
    public void ConicPaintAtCornerAndEdgeCentersRetainsFullCoverage(double centerX, double centerY) {
        var gradient = new OfficeConicGradient(centerX, centerY, 0D, new[] {
            new OfficeGradientStop(0D, OfficeColor.Red),
            new OfficeGradientStop(1D, OfficeColor.Blue)
        });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(gradient.CreateDrawing(80D, 60D));
        for (int y = 0; y < raster.Height; y++) {
            for (int x = 0; x < raster.Width; x++) Assert.Equal((byte)255, raster.GetPixel(x, y).A);
        }
    }
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ConicCoveragePreservesAuthoredPlateauEndpoints(bool lateTransition, bool varyingAlpha) {
        OfficeColor transitionColor = varyingAlpha
            ? OfficeColor.FromRgba(255, 0, 0, 0)
            : OfficeColor.Red;
        OfficeGradientStop[] stops = lateTransition
            ? new[] {
                new OfficeGradientStop(0D, OfficeColor.Blue),
                new OfficeGradientStop(0.98D, OfficeColor.Blue),
                new OfficeGradientStop(1D, transitionColor)
            }
            : new[] {
                new OfficeGradientStop(0D, transitionColor),
                new OfficeGradientStop(0.02D, OfficeColor.Blue),
                new OfficeGradientStop(1D, OfficeColor.Blue)
            };
        var gradient = new OfficeConicGradient(0.5D, 0.5D, 0D, stops);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(gradient.CreateDrawing(400D, 400D));

        // Each pixel lies more than two degrees inside the blue plateau,
        // away from its authored endpoint and any antialiased transition.
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(lateTransition ? 175 : 225, 50));
    }

}
