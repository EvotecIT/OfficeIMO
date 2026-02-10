using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Emphasis_Edge_Tests {
    [Fact]
    public void Triple_Closer_Can_Close_Italic_Then_Bold() {
        var md = "**bold *italic***";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<strong>bold <em>italic</em></strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Unclosed_Emphasis_Is_Literal() {
        var md = "*not closed";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("*not closed", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<em>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Intraword_Underscores_Do_Not_Create_Emphasis() {
        var md = "foo_bar_baz";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("foo_bar_baz", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<em>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Strikethrough_Can_Nest_Inside_Emphasis() {
        var md = "*a ~~b~~ c*";
        var html = MarkdownReader.Parse(md).ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<em>a <del>b</del> c</em>", html, StringComparison.Ordinal);
    }
}

