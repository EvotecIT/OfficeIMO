using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Theory]
    [InlineData("calc(20px + 10%)", 40D)]
    [InlineData("min(50px, 30%)", 50D)]
    [InlineData("max(20px, 10%)", 20D)]
    [InlineData("clamp(20px, 40%, 60px)", 60D)]
    [InlineData("calc((8px + 2px) * 3)", 30D)]
    [InlineData("calc(90px / 3)", 30D)]
    [InlineData("calc((10px / 2px) * 5px)", 25D)]
    public void HtmlCssLengthMath_ResolvesAcrossSharedLayoutConsumers(string width, double expected) {
        string html = "<div id='math' style='width:" + width + ";height:10px;margin:0;background:#ff0000'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        HtmlRenderShape shape = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#math");
        Assert.Equal(expected, shape.Width, 3);
    }

    [Theory]
    [InlineData("calc(10px + 2)")]
    [InlineData("calc(10px * 2px)")]
    [InlineData("calc(10px / 0)")]
    [InlineData("min(10px, 2)")]
    [InlineData("clamp(10px, 20px)")]
    [InlineData("calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(1px)))))))))))))))))))))))))))))))))")]
    public void HtmlCssLengthMath_RejectsInvalidDimensionsAndUnboundedNesting(string width) {
        string html = "<div id='math-invalid' style='width:" + width + ";height:10px;margin:0;background:#ff0000'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape shape = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#math-invalid");
        Assert.NotEqual(10D, shape.Width);
    }
}
