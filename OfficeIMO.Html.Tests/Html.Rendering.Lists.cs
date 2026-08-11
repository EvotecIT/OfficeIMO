using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_UsesCanonicalHtmlListOrdinals() {
        const string html = "<ol start='9x'><li>First</li><li value='12junk'>Second</li><li>Third</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { "9. ", "First", "12. ", "Second", "13. ", "Third" },
            rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("upper-roman", "IV. ")]
    [InlineData("lower-alpha", "d. ")]
    [InlineData("lower-greek", "δ. ")]
    [InlineData("decimal-leading-zero", "04. ")]
    [InlineData("cjk-decimal", "四. ")]
    [InlineData("hiragana", "え. ")]
    [InlineData("katakana", "エ. ")]
    public void HtmlRendering_FormatsStandardOrderedListCounterStyles(string style, string marker) {
        string html = "<ol start='4' style='list-style-type:" + style + "'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("cjk-decimal", 204, "二〇四. ")]
    [InlineData("full-width", 204, "２０４. ")]
    [InlineData("cjk-heavenly-stem", 10, "癸. ")]
    [InlineData("cjk-earthly-branch", 12, "亥. ")]
    public void HtmlRendering_FormatsBoundedEastAsianCounterStyles(string style, int start, string marker) {
        string html = "<ol start='" + start + "' style='list-style-type:" + style + "'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("circle", "◦ ")]
    [InlineData("square", "▪ ")]
    [InlineData("'→'", "→ ")]
    public void HtmlRendering_FormatsUnorderedAndQuotedListMarkers(string style, string marker) {
        string html = "<ul style=\"list-style-type:" + style + "\"><li>Item</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }
}
