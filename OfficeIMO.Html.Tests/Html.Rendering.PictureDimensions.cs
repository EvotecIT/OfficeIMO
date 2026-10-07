using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("width='320' height='160'", "", 320D, 160D)]
    [InlineData("width='320'", "", 320D, 160D)]
    [InlineData("height='160'", "", 320D, 160D)]
    [InlineData("", "", 80D, 40D)]
    [InlineData("width='' height=''", "", 4D, 2D)]
    [InlineData("width='320' height='160'", "width:120px;height:auto", 120D, 60D)]
    [InlineData("width='320' height='160'", "width:auto;height:auto", 4D, 2D)]
    public void HtmlRender_PictureUsesSelectedSourceDimensionHints(string sourceDimensions, string css, double width, double height) {
        byte[] png = PdfPngTestImages.CreateRgbPng(4, 2);
        string source = "data:image/png;base64," + Convert.ToBase64String(png);
        string html = $"<body style='margin:0'><picture><source media='(min-width:300px)' srcset='{source}' {sourceDimensions}>"
            + $"<img id='photo' src='{source}' width='80' height='40' style='display:block;{css}'></picture></body>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(width, image.Width, 5);
        Assert.Equal(height, image.Height, 5);
    }

    [Theory]
    [InlineData("media='(max-width:1px)'", false)]
    [InlineData("type='image/unsupported-test-format'", false)]
    [InlineData("", true)]
    public void HtmlRender_PictureIgnoresUnselectedSourceDimensions(string sourceCondition, bool afterImage) {
        byte[] png = PdfPngTestImages.CreateRgbPng(4, 2);
        string source = "data:image/png;base64," + Convert.ToBase64String(png);
        string sourceElement = $"<source {sourceCondition} srcset='{source}' width='320' height='160'>";
        string img = $"<img src='{source}' width='80' height='40' style='display:block'>";
        string html = "<body style='margin:0'><picture>" + (afterImage ? img + sourceElement : sourceElement + img) + "</picture></body>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(80D, image.Width, 5);
        Assert.Equal(40D, image.Height, 5);
    }
    [Theory]
    [InlineData("display:flex;align-items:flex-start")]
    [InlineData("display:grid;grid-template-columns:max-content 20px")]
    public void HtmlRender_PictureSourceHintsReachContainerSizing(string containerCss) {
        byte[] png = PdfPngTestImages.CreateRgbPng(4, 2);
        string source = "data:image/png;base64," + Convert.ToBase64String(png);
        string html = $"<body style='margin:0'><div style='{containerCss}'>"
            + $"<picture><source srcset='{source}' width='320' height='160'><img src='{source}' width='80' height='40' style='display:block'></picture>"
            + "<div id='after' style='width:20px;height:20px;background:blue'></div></div></body>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        HtmlRenderShape following = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#after");
        Assert.Equal(320D, image.Width, 5);
        Assert.Equal(160D, image.Height, 5);
        Assert.Equal(320D, following.X, 5);
    }

    [Theory]
    [InlineData("position:absolute;left:0;top:0", "height='160'", 320D, 160D)]
    [InlineData("position:fixed;left:0;top:0", "height='160'", 320D, 160D)]
    [InlineData("position:absolute;left:0;right:0;top:0;bottom:0", "width='320'", 320D, 160D)]
    [InlineData("position:fixed;left:0;right:0;top:0;bottom:0", "width='320'", 320D, 160D)]
    [InlineData("position:absolute;left:0;right:0;top:0;bottom:0", "width='' height=''", 4D, 2D)]
    [InlineData("position:fixed;left:0;right:0;top:0;bottom:0", "width='' height=''", 4D, 2D)]
    [InlineData("float:left", "height='160'", 320D, 160D)]
    public void HtmlRender_PicturePartialHintsPreserveIntrinsicPositionedSizing(string css, string dimensions, double width, double height) {
        byte[] png = PdfPngTestImages.CreateRgbPng(4, 2);
        string source = "data:image/png;base64," + Convert.ToBase64String(png);
        string html = $"<body style='margin:0'><picture><source srcset='{source}' {dimensions}>"
            + $"<img src='{source}' width='80' height='40' style='display:block;{css}'></picture></body>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 400D, ViewportHeight = 300D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(width, image.Width, 5);
        Assert.Equal(height, image.Height, 5);
    }

}
