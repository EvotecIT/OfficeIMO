using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_BrowserUserAgentStylesInsetAnUnstyledBody() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White
        };
        options.UseBrowserUserAgentStyles();

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<body><div style='width:10px;height:10px;background:#ff0000'></div></body>",
            options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(OfficeColor.White, raster.GetPixel(7, 7));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(8, 8));
        Assert.Equal("serif", options.DefaultFontFamily);
        Assert.Equal(HtmlRenderUserAgentStyleMode.Browser, options.UserAgentStyles);
        HtmlRenderOptions clone = options.Clone();
        Assert.Equal("serif", clone.DefaultFontFamily);
        Assert.Equal(HtmlRenderUserAgentStyleMode.Browser, clone.UserAgentStyles);
    }

    [Fact]
    public void HtmlRender_AuthoredBodyMarginOverridesBrowserUserAgentDefault() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White
        };
        options.UseBrowserUserAgentStyles();

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>body{margin:0}</style><body><div style='width:10px;height:10px;background:#ff0000'></div></body>",
            options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(0, 0));
    }
}
