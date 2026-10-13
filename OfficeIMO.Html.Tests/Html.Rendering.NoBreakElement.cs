using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("AAAAAA BBBBBB", "AAAAAA BBBBBB")]
    [InlineData("m<sup>−1</sup> · kg · s<sup>−<span style='font-family:Helvetica'>2</span></sup>", "m−1 · kg · s−2")]
    public void HtmlNobr_DefaultKeepsTheCompleteInlineRangeOnOneLine(string content, string expected) {
        HtmlRenderText[] text = RenderNobrFixture("<nobr>" + content + "</nobr>");

        Assert.Equal(expected, string.Concat(text.Select(item => item.Text)));
        Assert.Single(text.Select(item => item.LayoutY).Distinct());
    }

    [Theory]
    [InlineData("<nobr style='white-space:normal'>AAAAAA BBBBBB</nobr>")]
    [InlineData("<style>nobr{white-space:normal}</style><nobr>AAAAAA BBBBBB</nobr>")]
    public void HtmlNobr_LocalAuthorNormalAllowsWrapping(string content) {
        AssertNobrLines(content, wraps: true);
    }

    [Theory]
    [InlineData("<style>#ancestor{white-space:normal!important}</style><span id='ancestor'><nobr>AAAAAA BBBBBB</nobr></span>", false)]
    [InlineData("<style>#ancestor{white-space:nowrap!important}</style><span id='ancestor'><nobr style='white-space:normal'>AAAAAA BBBBBB</nobr></span>", true)]
    public void HtmlNobr_AncestorSpecificityDoesNotReplaceTheElementsOwnValue(string content, bool wraps) {
        AssertNobrLines(content, wraps);
    }

    [Theory]
    [InlineData("", true)]
    [InlineData("white-space:nowrap", false)]
    public void HtmlNobr_WbrDefaultAllowsABreakUnlessTheAuthorOverridesIt(string style, bool wraps) {
        AssertNobrLines("<nobr>AAAAAA<wbr style='" + style + "'>BBBBBB</nobr>", wraps);
    }

    [Fact]
    public void HtmlNobr_MinContentUsesTheWholeNoBreakRange() {
        double actual = Width("<nobr>AAAAAA BBBBBB</nobr>");
        double control = Width("<span style='white-space:nowrap'>AAAAAA BBBBBB</span>");

        Assert.True(actual > 60D);
        Assert.Equal(control, actual, 3);

        static double Width(string content) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
                "<span id='box' style='display:inline-block;width:min-content;background:red;font:16px/24px Courier'>"
                + content + "</span>"), NobrRenderOptions());
            return Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderShape>(),
                shape => shape.Source == "span#box").Width;
        }
    }

    // Separate filter: these check CSS-wide/default interaction without depending
    // on the existing paragraph-level nowrap handling of child-normal text.
    [Theory]
    [InlineData("initial", true)]
    [InlineData("inherit", true)]
    [InlineData("unset", true)]
    [InlineData("revert", false)]
    [InlineData("revert-layer", false)]
    public void HtmlNobrCssWide_ExplicitResetAndInheritanceRemainDistinctFromUaRollback(string keyword, bool wraps) {
        AssertNobrLines("<span style='white-space:normal'><nobr style='white-space:" + keyword
            + "'>AAAAAA BBBBBB</nobr></span>", wraps);
    }

    private static void AssertNobrLines(string content, bool wraps) {
        HtmlRenderText[] text = RenderNobrFixture(content);
        Assert.Equal("AAAAAABBBBBB", string.Concat(text.Select(item => item.Text)).Replace(" ", string.Empty));
        int lines = text.Select(item => item.LayoutY).Distinct().Count();
        if (wraps) Assert.True(lines > 1);
        else Assert.Equal(1, lines);
    }

    private static HtmlRenderText[] RenderNobrFixture(string content) =>
        EnumerateRenderVisuals(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<div style='width:60px;font:16px/24px Courier'>" + content + "</div>"), NobrRenderOptions())
            .Pages[0].Scene).OfType<HtmlRenderText>().ToArray();

    private static HtmlRenderOptions NobrRenderOptions() => new() {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 180D,
        ViewportHeight = 100D,
        Margins = HtmlRenderMargins.All(0D)
    };
}
