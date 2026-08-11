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
    [InlineData("japanese-informal", 10, "十、")]
    [InlineData("japanese-formal", 101, "壱百壱、")]
    [InlineData("korean-hangul-formal", 6001, "육천일, ")]
    [InlineData("korean-hanja-informal", 101, "百一, ")]
    [InlineData("korean-hanja-formal", 11, "壹拾壹, ")]
    [InlineData("simp-chinese-informal", 101, "一百零一、")]
    [InlineData("simp-chinese-formal", 6001, "陆仟零壹、")]
    [InlineData("trad-chinese-informal", 10, "十、")]
    [InlineData("trad-chinese-formal", 99, "玖拾玖、")]
    [InlineData("cjk-ideographic", 6001, "六千零一、")]
    public void HtmlRendering_FormatsLonghandEastAsianCounterStyles(string style, int start, string marker) {
        string html = "<ol start='" + start + "' style='list-style-type:" + style + "'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("japanese-informal", -11, "マイナス十一")]
    [InlineData("korean-hangul-formal", -11, "마이너스 일십일")]
    [InlineData("simp-chinese-formal", -101, "负壹佰零壹")]
    [InlineData("trad-chinese-formal", -101, "負壹佰零壹")]
    [InlineData("japanese-formal", 10000, "一〇〇〇〇")]
    public void HtmlRendering_FormatsLonghandEastAsianRepresentations(string style, int value, string expected) {
        Assert.True(HtmlCounterStyleFormatter.TryFormat(value, style, out string formatted));
        Assert.Equal(expected, formatted);
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

    [Fact]
    public void HtmlRendering_FormatsAuthorDefinedCounterStyleMarkers() {
        const string html = """
            <style>
              @counter-style binary {
                system:numeric;
                symbols:"0" "1";
                pad:4 "0";
                prefix:"[";
                suffix:"] ";
              }
              ol { list-style-type:binary; }
            </style>
            <ol start="3"><li>Item</li></ol>
            """;

        var options = new HtmlRenderOptions();
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        HtmlCounterStyleRegistry registry = HtmlCounterStyleRegistry.Parse(document, options);
        Assert.True(registry.TryFormatMarker(3, "binary", out string directMarker));
        Assert.Equal("[0011] ", directMarker);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(new[] { "[0011] ", "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_AppliesCounterStyleConditionalRulesForTheActiveMedia() {
        const string html = """
            <style>
              @media screen { @counter-style screen-only { system:cyclic; symbols:"S"; } }
              @media print { @counter-style print-only { system:cyclic; symbols:"P"; } }
            </style>
            """;
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();

        HtmlCounterStyleRegistry registry = HtmlCounterStyleRegistry.Parse(document, options);

        Assert.True(registry.TryFormat(1, "print-only", out string printed));
        Assert.Equal("P", printed);
        Assert.False(registry.TryFormat(1, "screen-only", out _));
    }

    [Theory]
    [InlineData("<ol start='2147483647' style='list-style-type:symbols(symbolic &quot;x&quot;)'><li>Item</li></ol>")]
    [InlineData("<style>@counter-style huge{system:symbolic;symbols:'x'}</style><ol start='2147483647' style='list-style-type:huge'><li>Item</li></ol>")]
    [InlineData("<style>@counter-style huge{system:additive;additive-symbols:1 'x'}</style><ol start='2147483647' style='list-style-type:huge'><li>Item</li></ol>")]
    public void HtmlRendering_BoundsExpandedCounterRepresentationsAndUsesDecimalFallback(string html) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { "2147483647. ", "Item" }, rendered.Text.Split('\n'));
        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.CounterRepresentationLimitExceeded);
    }

}
