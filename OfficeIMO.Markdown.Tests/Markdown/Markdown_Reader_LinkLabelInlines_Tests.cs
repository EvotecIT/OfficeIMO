using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_LinkLabelInlines_Tests {
    [Fact]
    public void Link_Label_Can_Contain_Emphasis() {
        var md = "[*x*](https://example.com)";
        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md);

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
        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<a href=\"https://example.com\"><em>x</em></a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Strong_Can_Wrap_Inline_Link() {
        const string md = "**[Installation](/docs/installation/)**";
        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md);

        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<strong><a href=\"/docs/installation/\">Installation</a></strong>", html, StringComparison.Ordinal);

        var round = doc.ToMarkdown().Trim();
        Assert.Equal(md, round);
    }

    [Fact]
    public void Strong_Can_Wrap_Reference_Link() {
        const string md = """
**[Installation][install]**

[install]: /docs/installation/
""";
        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md);

        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<strong><a href=\"/docs/installation/\">Installation</a></strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Link_Destination_With_Source_Line_Break_Remains_Text() {
        const string md = """
[link](<foo
bar>)
""";

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, MarkdownReaderOptions.CreateCommonMarkProfile());
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));

        Assert.Empty(paragraph.Inlines.Nodes.OfType<LinkInline>());
        var rawHtml = Assert.Single(paragraph.Inlines.Nodes.OfType<HtmlRawInline>());
        Assert.Equal("<foo\nbar>", rawHtml.Html);

        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("[link](<foo\nbar>)", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Link_Label_Does_Not_Skip_Literal_Html_When_InlineHtml_Is_Disabled() {
        const string md = "[<u>[</u>](url)";
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.InlineHtml = false;

        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md, options);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        var link = Assert.Single(paragraph.Inlines.Nodes.OfType<LinkInline>());
        Assert.Equal("</u>", link.Text);
        Assert.Contains("[&lt;u&gt;", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"url\">&lt;/u&gt;</a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<a href=\"url\">&lt;u&gt;[&lt;/u&gt;</a>", html, StringComparison.Ordinal);
    }
}

