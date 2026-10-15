using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlOutlinePaintTests {
    [Theory]
    [InlineData(HtmlRenderMode.Paged)]
    [InlineData(HtmlRenderMode.Continuous)]
    public void OutlinePaintsAboveFollowingBlockBackgroundWithoutChangingFlowHeight(HtmlRenderMode mode) {
        HtmlRenderPage page = Assert.Single(HtmlRenderTestDriver.Render("<!doctype html><style>"
            + "html,body{margin:0;padding:0;font-size:0;line-height:0;background:white}"
            + "body{padding:16px}</style>"
            + "<div id='outlined' style='width:100px;height:20px;background:white;outline:2px solid blue'></div>"
            + "<div id='tail' style='width:120px;height:2px;background:#d06c24'></div>",
            new HtmlRenderOptions {
                Mode = mode,
                ViewportWidth = 240D, ViewportHeight = 180D,
                Margins = HtmlRenderMargins.All(0D),
                UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
                AllowSystemFontFallback = false
            }).Pages);
        HtmlRenderShape outlined = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source == "div#outlined" && shape.Shape.FillColor.HasValue);
        HtmlRenderShape tail = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source == "div#tail" && shape.Shape.FillColor.HasValue);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        Assert.Equal(16D, outlined.Y, 6);
        Assert.Equal(20D, outlined.Height, 6);
        Assert.Equal(36D, tail.Y, 6);
        Assert.Equal(2D, tail.Height, 6);
        Assert.Equal(OfficeColor.White, raster.GetPixel(32, 24));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(32, 36));
        Assert.Equal(OfficeColor.FromRgb(208, 108, 36), raster.GetPixel(128, 36));
    }
}
