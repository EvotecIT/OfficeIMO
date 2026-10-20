using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("line-height:calc(10cqw - 20px)", 16D)]
    [InlineData("line-height:calc(10vw - 20px)", 16D)]
    [InlineData("font:16px/calc(10cqw - 20px) Pinned", 16D)]
    [InlineData("line-height:calc(20px - 10cqw)", 0D)]
    [InlineData("font:16px/calc(20px - 10cqw) Pinned", 0D)]
    public void HtmlFontMathRanges_MarginLeadingUsesTheActualPaperBeforeRangeClamping(string declaration, double expected) {
        var options = PageMarginTypographyOptions();
        options.PageSize = new OfficePageSize(360D / 96D, 480D / 96D);
        string html = PageMarginTypographySource("font:16px/22px Pinned;" + declaration, "Hgyp");

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options).Pages);

        Assert.Equal(expected, PageMarginTypographyText(page).LineHeight, 6);
    }

    [Theory]
    [InlineData("calc(10cqw - 24px)", 12D)]
    [InlineData("calc(36px - 10cqw)", 0D)]
    [InlineData("calc(20px - 10cqw)", 0D)]
    [InlineData("0", 0D)]
    public void HtmlFontMathRanges_ZeroMarginSizeRemainsZeroAndDoesNotPaint(string size, double expected) {
        var options = PageMarginTypographyOptions();
        options.PageSize = new OfficePageSize(360D / 96D, 480D / 96D);
        string html = PageMarginTypographySource("font:16px/22px Pinned;font-size:" + size, "Hgyp");
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        if (expected == 0D) {
            Assert.DoesNotContain(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(), text => text.Source == "@page @bottom-center");
            Assert.False(rendered.HasLoss);
        } else Assert.Equal(expected, PageMarginTypographyText(page).Font.Size, 6);
    }

    [Theory]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser)]
    [InlineData(HtmlRenderUserAgentStyleMode.Document)]
    public void HtmlFontMathRanges_OrdinaryTextUsesTheSameComputedRangePolicy(HtmlRenderUserAgentStyleMode mode) {
        var options = NaturalLineOptions();
        options.UserAgentStyles = mode;
        const string html = """
            <style>body,p{margin:0;font:16px Pinned}
              #subject{font-size:calc(10vw - 20px);line-height:calc(10vw - 20px)}
              #hidden{font-size:calc(20px - 10vw)}
            </style><p id='subject'>Visible</p><p id='hidden'>Hidden</p>
            """;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>().ToArray();

        HtmlRenderText visible = Assert.Single(text, item => item.Text == "Visible");
        Assert.Equal(60D, visible.Font.Size, 6);
        Assert.Equal(60D, visible.LineHeight, 6);
        Assert.DoesNotContain(text, item => item.Text == "Hidden");
    }
}
