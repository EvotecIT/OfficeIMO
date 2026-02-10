using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_LinkLabelInlines_Tests {
    [Fact]
    public void Link_Label_Can_Contain_Emphasis() {
        var md = "[*x*](https://example.com)";
        var doc = MarkdownReader.Parse(md);

        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<a href=\"https://example.com\"><em>x</em></a>", html, StringComparison.Ordinal);

        var round = doc.ToMarkdown().Trim();
        Assert.Equal(md, round);
    }

    [Fact]
    public void Reference_Link_Label_Can_Contain_Inline_Markup() {
        var md = """
[*x*][r]

[r]: https://example.com
""";
        var doc = MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<a href=\"https://example.com\"><em>x</em></a>", html, StringComparison.Ordinal);
    }
}

