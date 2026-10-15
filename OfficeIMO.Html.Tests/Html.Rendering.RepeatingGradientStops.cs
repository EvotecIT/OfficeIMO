using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRepeatingGradient_KeepsEveryStripeSolidBeyondTheFirstCycles() {
        const string html = "<div style='width:180px;height:20px;background-color:red;background-image:repeating-linear-gradient(to right,rgba(255 0 0 / 50%) 0 10px,color-mix(in srgb,blue 75%,white) 10px 20px)'></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderShape shape = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            candidate => candidate.Shape.FillGradient != null);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        for (int stripe = 0; stripe < 18; stripe++) {
            OfficeColor expected = stripe % 2 == 0 ? OfficeColor.Red : OfficeColor.FromRgb(64, 64, 255);
            Assert.Equal(expected, raster.GetPixel((int)shape.X + stripe * 10 + 5, (int)shape.Y + 10));
        }
    }

    [Theory]
    [InlineData(180D, 0D)]
    [InlineData(214D, 7D)]
    [InlineData(214D, -7D)]
    public void HtmlRepeatingGradient_PreservesOrderedCoincidentStopsAtEveryCycleBoundary(double referenceLength, double firstPixels) {
        string first = firstPixels.ToString(System.Globalization.CultureInfo.InvariantCulture);
        string middle = (firstPixels + 10D).ToString(System.Globalization.CultureInfo.InvariantCulture);
        string last = (firstPixels + 20D).ToString(System.Globalization.CultureInfo.InvariantCulture);
        Assert.True(HtmlCssGradientStops.TryParse(
            new[] { "red " + first + "px " + middle + "px", "blue " + middle + "px " + last + "px" },
            startIndex: 0, maximumStops: 128, out HtmlCssGradientStops? definition, out _));
        Assert.True(definition!.TryResolve(referenceLength, 16D, 16D, referenceLength, 40D,
            repeating: true, out IReadOnlyList<OfficeGradientStop>? stops, out bool stopLimitExceeded));

        Assert.False(stopLimitExceeded);
        AssertRepeatingStripeBoundaries(stops!, referenceLength, firstPixels, 10D);
    }

    [Fact]
    public void HtmlRepeatingGradient_PreservesCoincidentAngularCycleBoundaries() {
        Assert.True(HtmlCssGradientStops.TryParseConic(
            new[] { "red 0deg 20deg", "blue 20deg 40deg" },
            startIndex: 0, maximumStops: 128, out HtmlCssGradientStops? definition, out _));
        Assert.True(definition!.TryResolveConic(repeating: true,
            out IReadOnlyList<OfficeGradientStop>? stops, out bool stopLimitExceeded));

        Assert.False(stopLimitExceeded);
        AssertRepeatingStripeBoundaries(stops!, 360D, 0D, 20D);
    }

    [Fact]
    public void HtmlRepeatingGradient_PreservesSmoothCycleIntervals() {
        Assert.True(HtmlCssGradientStops.TryParse(new[] { "red 0px", "blue 20px" },
            startIndex: 0, maximumStops: 128, out HtmlCssGradientStops? definition, out _));
        Assert.True(definition!.TryResolve(180D, 16D, 16D, 180D, 40D,
            repeating: true, out IReadOnlyList<OfficeGradientStop>? stops, out _));

        for (int cycle = 0; cycle < 9; cycle++) {
            OfficeGradientStop[] interval = stops!.Where(stop =>
                stop.Offset >= (cycle + .25D) / 9D && stop.Offset <= (cycle + .75D) / 9D).ToArray();
            Assert.Empty(interval);
        }
        AssertRepeatingStripeBoundaries(stops!, 180D, 0D, 20D, cycleBoundariesOnly: true);
    }

    [Fact]
    public void HtmlRepeatingGradient_ZeroPeriodUsesTheLastColor() {
        Assert.True(HtmlCssGradientStops.TryParse(new[] { "red 10px", "blue 10px" },
            startIndex: 0, maximumStops: 128, out HtmlCssGradientStops? definition, out _));
        Assert.True(definition!.TryResolve(180D, 16D, 16D, 180D, 40D,
            repeating: true, out IReadOnlyList<OfficeGradientStop>? stops, out bool stopLimitExceeded));

        Assert.False(stopLimitExceeded);
        Assert.Equal(new[] { new OfficeGradientStop(0D, OfficeColor.Blue), new OfficeGradientStop(1D, OfficeColor.Blue) }, stops);
    }

    private static void AssertRepeatingStripeBoundaries(IReadOnlyList<OfficeGradientStop> stops,
        double referenceLength, double first, double stripeWidth, bool cycleBoundariesOnly = false) {
        Assert.Equal(0D, stops[0].Offset);
        Assert.Equal(1D, stops[stops.Count - 1].Offset);
        int firstBoundary = (int)Math.Ceiling(-first / stripeWidth);
        int lastBoundary = (int)Math.Floor((referenceLength - first) / stripeWidth);
        for (int boundary = firstBoundary; boundary <= lastBoundary; boundary++) {
            double expectedOffset = (first + boundary * stripeWidth) / referenceLength;
            if (expectedOffset <= 0D || expectedOffset >= 1D) continue;
            OfficeGradientStop[] pair = stops.Where(stop => Math.Abs(stop.Offset - expectedOffset) < 0.000000000001D).ToArray();
            Assert.Equal(2, pair.Length);
            Assert.Equal(pair[0].Offset, pair[1].Offset);
            bool startsRed = cycleBoundariesOnly || boundary % 2 == 0;
            Assert.Equal(startsRed ? OfficeColor.Blue : OfficeColor.Red, pair[0].Color);
            Assert.Equal(startsRed ? OfficeColor.Red : OfficeColor.Blue, pair[1].Color);
        }
    }
}
