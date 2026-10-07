using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgTextViewportTests {
    [Theory]
    [InlineData(-12, 20, 8, "")]
    [InlineData(100, -92, 8, "")]
    [InlineData(0, 76, 76, "")]
    [InlineData(-12, 20, 8, "font-style='italic'")]
    [InlineData(-12, 20, 8, "font-weight='bold'")]
    public void TranslatedSvgTextIsClippedAtItsDestinationViewport(int localX, int translationX, int destinationX, string style) {
        OfficeDrawing translated = Read("<g transform='translate(" + translationX + " 30)'><text x='" + localX + "' y='0' " + style + ">A</text></g>");
        OfficeDrawing direct = Read("<text x='" + destinationX + "' y='30' " + style + ">A</text>");

        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(direct);
        Assert.Contains(Enumerable.Range(0, expected.Width * expected.Height), index => expected.GetPixel(index % expected.Width, index / expected.Width).A > 0);
        byte[] actualPixels = OfficeDrawingRasterRenderer.Render(translated).GetPixels();
        byte[] expectedPixels = expected.GetPixels();
        int maximumDifference = expectedPixels.Zip(actualPixels, (expectedValue, actualValue) => System.Math.Abs(expectedValue - actualValue)).Max();
        // Synthetic styles use fractional contours; one-channel rounding can differ
        // when the same ink is composited from a translated intermediate surface.
        Assert.True(maximumDifference <= (style.Length == 0 ? 0 : 1), "Maximum channel difference: " + maximumDifference);
        Assert.Contains(">A</text>", OfficeDrawingSvgExporter.ToSvg(translated), System.StringComparison.Ordinal);
        if (destinationX == 76) Assert.Contains("<clipPath", OfficeDrawingSvgExporter.ToSvg(translated), System.StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("scale(.5)")]
    [InlineData("rotate(15 40 20)")]
    [InlineData("skewX(15)")]
    public void NestedAffineSvgTextRetainsItsCompleteLocalGlyphs(string transform) {
        OfficeDrawing translated = Read("<g transform='" + transform + "'><g transform='translate(20 30)'><text x='-12' y='0'>A</text></g></g>");
        OfficeDrawing direct = Read("<g transform='" + transform + "'><text x='8' y='30'>A</text></g>");

        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(direct);
        Assert.Contains(Enumerable.Range(0, expected.Width * expected.Height), index => expected.GetPixel(index % expected.Width, index / expected.Width).A > 0);
        Assert.Equal(expected.GetPixels(), OfficeDrawingRasterRenderer.Render(translated).GetPixels());
    }

    [Fact]
    public void FittedRootViewportKeepsTranslatedTextAndRegisteredFontPaint() {
        const string viewBox = "viewBox='0 0 40 40'";
        OfficeDrawing translated = Read("<g transform='translate(-92 30)'><text x='100' y='0'>A</text></g>", viewBox);
        OfficeDrawing direct = Read("<text x='8' y='30'>A</text>", viewBox);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(direct).GetPixels(), OfficeDrawingRasterRenderer.Render(translated).GetPixels());
    }

    private static OfficeDrawing Read(string content, string viewport = "") {
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("FixtureFont", ManagedTextShapingTestAssets.CreateFont('A'));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='80' height='40' font-family='FixtureFont' font-size='20' " + viewport + ">" + content + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        return drawing!;
    }
}
