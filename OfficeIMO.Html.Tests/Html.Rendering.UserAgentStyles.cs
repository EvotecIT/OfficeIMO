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

    [Theory]
    [InlineData("margin:initial")]
    [InlineData("margin:unset")]
    [InlineData("margin:inherit")]
    [InlineData("all:initial")]
    [InlineData("all:unset")]
    public void HtmlRender_CssWideInitialBodyResetsOverrideBrowserUserAgentMargin(string declaration) {
        HtmlRenderOptions options = BrowserOptions();

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>body{" + declaration + "}</style><body><div style='width:10px;height:10px;background:#ff0000'></div></body>",
            options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(0, 0));
    }

    [Theory]
    [InlineData("margin:revert")]
    [InlineData("margin:revert-layer")]
    [InlineData("all:revert")]
    public void HtmlRender_CssWideRevertExposesBrowserUserAgentMargin(string declaration) {
        HtmlRenderOptions options = BrowserOptions();

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>body{" + declaration + "}</style><body><div style='width:10px;height:10px;background:#ff0000'></div></body>",
            options);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());

        Assert.Equal(OfficeColor.White, raster.GetPixel(7, 7));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(8, 8));
    }

    [Theory]
    [InlineData("horizontal-tb", "ltr", "margin-inline-start:initial", 8D, 8D, 8D, 0D)]
    [InlineData("horizontal-tb", "rtl", "margin-inline-start:unset", 8D, 0D, 8D, 8D)]
    [InlineData("horizontal-tb", "ltr", "margin-inline:initial", 8D, 0D, 8D, 0D)]
    [InlineData("horizontal-tb", "ltr", "margin-block:unset", 0D, 8D, 0D, 8D)]
    [InlineData("vertical-rl", "ltr", "margin-inline-start:initial", 0D, 8D, 8D, 8D)]
    [InlineData("vertical-rl", "ltr", "margin-block-start:unset", 8D, 0D, 8D, 8D)]
    [InlineData("vertical-lr", "rtl", "margin-inline-start:initial", 8D, 8D, 0D, 8D)]
    [InlineData("sideways-lr", "ltr", "margin-block-start:unset", 8D, 8D, 8D, 0D)]
    [InlineData("horizontal-tb", "ltr", "margin-left:20px;margin-inline-start:initial", 8D, 8D, 8D, 0D)]
    [InlineData("horizontal-tb", "ltr", "margin-inline-start:initial;margin-left:20px", 8D, 8D, 8D, 20D)]
    [InlineData("horizontal-tb", "ltr", "margin-inline-start:revert", 8D, 8D, 8D, 8D)]
    [InlineData("horizontal-tb", "ltr", "margin-inline-start:revert-layer", 8D, 8D, 8D, 8D)]
    public void HtmlRender_LogicalCssWideBodyMarginsMapResetProvenanceToPhysicalSides(
        string writingMode,
        string direction,
        string declaration,
        double expectedTop,
        double expectedRight,
        double expectedBottom,
        double expectedLeft) {
        string html = "<body style='writing-mode:" + writingMode + ";direction:" + direction + ";"
            + declaration + "'></body>";
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(document);
        var styles = new HtmlComputedStyleSet(computed, new Dictionary<AngleSharp.Dom.IElement, HtmlPseudoElementStylePair>());
        HtmlRenderBoxStyle body = new HtmlRenderStyleResolver(styles, BrowserOptions(), new HtmlDiagnosticReport())
            .Resolve(document.Body!, 40D);

        Assert.Equal(expectedTop, body.MarginTop, 3);
        Assert.Equal(expectedRight, body.MarginRight, 3);
        Assert.Equal(expectedBottom, body.MarginBottom, 3);
        Assert.Equal(expectedLeft, body.MarginLeft, 3);
    }

    private static HtmlRenderOptions BrowserOptions() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White
        };
        options.UseBrowserUserAgentStyles();
        return options;
    }
}
