using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_LinkLabelInlines_Tests {
    public static IEnumerable<object[]> BareAutolinkLabels() {
        foreach (string label in new[] { "https://example.com/guide", "www.example.com", "person@example.com", "ftp://example.com/file" }) {
            yield return new object[] { "[" + label + "](https://outer.example)", label };
            yield return new object[] { "[" + label + "][ref]\n\n[ref]: https://outer.example", label };
            yield return new object[] { "[" + label + "][]\n\n[" + label + "]: https://outer.example", label };
            yield return new object[] { "[" + label + "]\n\n[" + label + "]: https://outer.example", label };
        }
    }

    [Theory]
    [MemberData(nameof(BareAutolinkLabels))]
    public void BareAutolinksStayTextInsideEveryKindOfLinkLabel(string source, string label) {
        var options = new MarkdownReaderOptions { AutolinkBareSchemeUrls = true };
        var document = MarkdownReader.Parse(source, options);
        var html = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = System.Xml.Linq.XElement.Parse(html.Trim());
        var link = Assert.Single(paragraph.Descendants("a"));
        Assert.Equal("https://outer.example", (string?)link.Attribute("href"));
        Assert.Equal(label, link.Value);
        Assert.Empty(link.Descendants("a"));

        var standalone = System.Xml.Linq.XElement.Parse(MarkdownReader.Parse(label, options).ToHtmlFragment().Trim());
        Assert.Single(standalone.Descendants("a"));
    }

    [Fact]
    public void FormattingInsideLinkLabelsDoesNotReenableAutolinks() {
        const string source = "[**https://example.com** <u>person@example.com</u>](https://outer.example)";
        var paragraph = System.Xml.Linq.XElement.Parse(MarkdownReader.Parse(source).ToHtmlFragment().Trim());
        var link = Assert.Single(paragraph.Descendants("a"));
        Assert.Equal("https://outer.example", (string?)link.Attribute("href"));
        Assert.Equal("https://example.com", Assert.Single(link.Descendants("strong")).Value);
        Assert.Equal("person@example.com", Assert.Single(link.Descendants("u")).Value);
    }

    [Theory]
    [InlineData("[<https://inner.example>](https://outer.example)", 1)]
    [InlineData("[<https://inner.example>][ref]\n\n[ref]: https://outer.example", 2)]
    [InlineData("[<u><https://inner.example></u>](https://outer.example)", 1)]
    [InlineData("[<sup><u><https://inner.example></u></sup>](https://outer.example)", 1)]
    [InlineData("[<ins><https://inner.example></ins>][ref]\n\n[ref]: https://outer.example", 2)]
    public void AngleAutolinkDeactivatesTheContainingLink(string source, int expectedLinks) {
        var document = MarkdownReader.Parse(source, MarkdownReaderOptions.CreateCommonMarkProfile());
        var html = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = System.Xml.Linq.XElement.Parse(html.Trim());
        var links = paragraph.Descendants("a").ToArray();
        Assert.Equal(expectedLinks, links.Length); // [ref] remains an independent shortcut reference.
        Assert.All(links, element => Assert.Empty(element.Descendants("a")));
        var link = links[0];
        Assert.Equal("https://inner.example", (string?)link.Attribute("href"));
        Assert.Equal("https://inner.example", link.Value);
    }

    [Theory]
    [InlineData("`<https://inner.example>`", false, true)]
    [InlineData("\\<https://inner.example>", false, true)]
    [InlineData("![<https://inner.example>](image.png)", true, true)]
    [InlineData("shown ![<https://inner.example>](image.png)", true, false)]
    public void CodeEscapesAndImageAltDoNotDeactivateTheOuterLink(string content, bool image, bool wrapper) {
        string label = wrapper ? "<u>" + content + "</u>" : content;
        var document = MarkdownReader.Parse("[" + label + "](https://outer.example)", MarkdownReaderOptions.CreateCommonMarkProfile());
        var paragraph = System.Xml.Linq.XElement.Parse(document.ToHtmlFragment().Trim());
        var link = Assert.Single(paragraph.Descendants("a"));
        Assert.Equal("https://outer.example", (string?)link.Attribute("href"));
        Assert.Equal(image ? 1 : 0, link.Descendants("img").Count());
    }

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
