using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlExpandedCornerTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(20, false)]
    [InlineData(0, true)]
    [InlineData(20, true)]
    public void ExpandedPaintPreservesSquareCornersWithoutAdvancingFlow(int height, bool shadow) {
        HtmlRenderPage page = Render(height, "0", shadow);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());
        HtmlRenderShape tail = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source == "div#tail" && shape.Shape.FillColor.HasValue);

        // Native zero-blur controls make these square-corner pixels fully covered.
        Assert.Equal(shadow ? OfficeColor.Red : OfficeColor.Blue,
            raster.GetPixel(shadow ? 11 : 14, shadow ? 11 : 14));
        Assert.Equal(16D + height, tail.Y, 6);
    }

    [Theory]
    [InlineData("0 8px 0 8px", false)]
    [InlineData("0 8px 0 8px", true)]
    [InlineData("0 8px 0 8px / 8px 8px 8px 8px", false)]
    [InlineData("0 8px 0 8px / 8px 8px 8px 8px", true)]
    [InlineData("8px 8px 8px 8px / 0 8px 0 8px", false)]
    [InlineData("8px 8px 8px 8px / 0 8px 0 8px", true)]
    public void ExpandedMixedCornersKeepSquareAndRoundedCoverage(string radii, bool shadow) {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Render(20, radii, shadow).CreateDrawing());

        // A zero radius on either axis keeps that corner square. Adjacent authored
        // rounded corners must remain rounded through the same expansion.
        Assert.Equal(shadow ? OfficeColor.Red : OfficeColor.Blue,
            raster.GetPixel(shadow ? 11 : 14, shadow ? 11 : 14));
        Assert.Equal(shadow ? OfficeColor.Red : OfficeColor.Blue,
            raster.GetPixel(shadow ? 100 : 97, shadow ? 40 : 37));
        Assert.Equal(OfficeColor.White, raster.GetPixel(shadow ? 100 : 97, shadow ? 11 : 14));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExpandedRoundedCornersRetainTheirAuthoredCurve(bool shadow) {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Render(20, "8px", shadow).CreateDrawing());

        Assert.Equal(OfficeColor.White, raster.GetPixel(shadow ? 11 : 14, shadow ? 11 : 14));
        Assert.Equal(shadow ? OfficeColor.Red : OfficeColor.Blue, raster.GetPixel(40, shadow ? 11 : 14));
    }

    [Fact]
    public void NegativeSpreadKeepsSquareCornersAndContractsTheShadow() {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(
            Render(20, "0", shadow: true, paint: "box-shadow:20px 0 0 -3px red").CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(112, 19));
        Assert.Equal(OfficeColor.White, raster.GetPixel(113, 19));
        Assert.Equal(OfficeColor.White, raster.GetPixel(112, 18));
    }

    private static HtmlRenderPage Render(int height, string radii, bool shadow, string? paint = null) =>
        Assert.Single(HtmlRenderTestDriver.Render("<!doctype html><style>"
            + "html,body{margin:0;padding:0;font-size:0;line-height:0;background:white}"
            + "body{padding:16px}</style>"
            + "<div id='box' style='width:80px;height:" + height + "px;background:white;border-radius:" + radii
            + ";" + (paint ?? (shadow ? "box-shadow:0 0 0 5px red" : "outline:2px solid blue")) + "'></div>"
            + "<div id='tail' style='width:120px;height:2px;background:#d06c24'></div>", new HtmlRenderOptions {
                ViewportWidth = 240D, ViewportHeight = 180D,
                Margins = HtmlRenderMargins.All(0D),
                UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
                AllowSystemFontFallback = false
            }).Pages);
}
