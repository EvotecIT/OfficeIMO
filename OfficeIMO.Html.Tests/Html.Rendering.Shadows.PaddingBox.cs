using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("0")]
    [InlineData("12px")]
    public void HtmlShadows_BorderedInsetStartsAtThePaddingEdge(string radii) {
        HtmlRenderPage page = RenderBorderedInsetBox("4px solid blue", radii);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        Assert.Equal(OfficeColor.Blue, raster.GetPixel(70, 35));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 36));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 38));
        Assert.Equal(OfficeColor.White, raster.GetPixel(70, 39));
        Assert.Equal(80D, Assert.Single(page.Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#tail").Y, 6);
    }

    [Fact]
    public void HtmlShadows_NegativeInsetSpreadUsesPaddingOriginAndRadii() {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(RenderBorderedInsetBox(
            "4px solid blue", "2px", "20px 20px 0 -8px red inset").CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 46));
        Assert.Equal(OfficeColor.White, raster.GetPixel(70, 48));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(47, 47));
        // Subtracting the 4px border makes the authored 2px padding radius zero.
        Assert.Equal(OfficeColor.White, raster.GetPixel(48, 48));
    }

    [Fact]
    public void HtmlShadows_AsymmetricBordersPreserveInsetBandWidths() {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(RenderBorderedInsetBox(
            "solid blue;border-width:2px 6px 4px 8px", "0").CreateDrawing());

        Assert.Equal(OfficeColor.Blue, raster.GetPixel(70, 33));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 34));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 36));
        Assert.Equal(OfficeColor.White, raster.GetPixel(70, 37));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(39, 50));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(40, 50));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(42, 50));
        Assert.Equal(OfficeColor.White, raster.GetPixel(43, 50));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(118, 50));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(120, 50));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 73));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(70, 74));
    }

    [Fact]
    public void HtmlShadows_ZeroBorderInsetRetainsItsSpreadGeometry() {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(
            RenderBorderedInsetBox("0", "12px").CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 32));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 34));
        Assert.Equal(OfficeColor.White, raster.GetPixel(70, 35));
    }

    [Theory]
    [InlineData("before")]
    [InlineData("after")]
    public void HtmlShadows_GeneratedInsetUsesThePaddingBox(string pseudo) {
        HtmlRenderPage page = RenderInsetPaddingBox("<div id='host'></div>",
            "#host::" + pseudo + "{content:'';display:block;width:80px;height:40px;"
            + "border:4px solid blue;background:white;box-shadow:0 0 0 3px red inset}");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        Assert.Equal(OfficeColor.Blue, raster.GetPixel(70, 35));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 36));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 38));
        Assert.Equal(OfficeColor.White, raster.GetPixel(70, 39));
    }

    [Fact]
    public void HtmlShadows_TableRowInsetUsesItsEffectiveZeroBorder() {
        HtmlRenderPage page = RenderInsetPaddingBox(
            "<table><tbody><tr id='row'><td></td></tr></tbody></table>",
            "table{border-collapse:separate;border-spacing:0}td{padding:0;width:80px;height:40px}"
            + "#row{border:4px solid blue;background:white;box-shadow:0 0 0 3px red inset}");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        // The table-row paint owner supplies a zero border override because
        // separate table rows do not paint their authored border declarations.
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 32));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(70, 34));
        Assert.Equal(OfficeColor.White, raster.GetPixel(70, 35));
    }

    [Theory]
    [InlineData(0, 40, 34, 40)]
    [InlineData(80, 0, 70, 34)]
    public void HtmlShadows_EmptyPaddingBoxDoesNotPaintThroughTransparentBorders(
        int width, int height, int x, int y) {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(RenderBorderedInsetBox(
            "4px solid transparent", "0", width: width, height: height).CreateDrawing());

        Assert.Equal(OfficeColor.White, raster.GetPixel(x, y));
    }

    private static HtmlRenderPage RenderBorderedInsetBox(
        string border, string radii, string shadow = "0 0 0 3px red inset", int width = 80, int height = 40) =>
        RenderInsetPaddingBox("<div id='box' style='width:" + width + "px;height:" + height
            + "px;border:" + border + ";border-radius:" + radii + ";background:white;box-shadow:"
            + shadow + "'></div><div id='tail' style='width:80px;height:2px;background:blue'></div>");

    private static HtmlRenderPage RenderInsetPaddingBox(string body, string css = "") =>
        Assert.Single(HtmlRenderTestDriver.Render("<!doctype html><style>"
            + "html,body{margin:0;padding:0;font-size:0;line-height:0;background:white}"
            + "body{padding:32px}" + css + "</style>" + body, new HtmlRenderOptions {
                ViewportWidth = 240D, ViewportHeight = 180D,
                Margins = HtmlRenderMargins.All(0D),
                UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
                AllowSystemFontFallback = false
            }).Pages);
}
