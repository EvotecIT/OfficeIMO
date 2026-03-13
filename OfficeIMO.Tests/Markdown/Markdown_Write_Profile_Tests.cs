using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Write_Profile_Tests {
    [Fact]
    public void Portable_Write_Profile_Degrades_Callouts_To_Quoted_Markdown() {
        var doc = MarkdownReader.Parse("""
> [!NOTE] Example
> Body text
""");

        var markdown = doc.ToMarkdown(MarkdownWriteOptions.CreatePortableProfile()).Replace("\r\n", "\n");

        Assert.DoesNotContain("[!NOTE]", markdown, StringComparison.Ordinal);
        Assert.Contains("> **Example**", markdown, StringComparison.Ordinal);
        Assert.Contains("> Body text", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Portable_Callout_Html_Fallback_Removes_OfficeImo_Callout_Chrome() {
        var doc = MarkdownReader.Parse("""
> [!NOTE] Example
> Body text
""");
        var options = new HtmlOptions { Kind = HtmlKind.Fragment, BodyClass = null };
        MarkdownBlockRenderBuiltInExtensions.AddPortableCalloutHtmlFallback(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<blockquote>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<strong>Example</strong>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Body text</p>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"callout", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Portable_Html_Fallbacks_Render_Toc_As_Plain_List() {
        var doc = MarkdownDoc.Create()
            .H2("Section")
            .H3("Child")
            .TocHere(options => {
                options.IncludeTitle = true;
                options.Title = "Contents";
                options.TitleLevel = 2;
                options.Layout = TocLayout.Panel;
            });
        var options = new HtmlOptions { Kind = HtmlKind.Fragment, BodyClass = null };
        MarkdownBlockRenderBuiltInExtensions.AddPortableHtmlFallbacks(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<h2>Contents</h2>", html, StringComparison.Ordinal);
        Assert.Contains("<ul>", html, StringComparison.Ordinal);
        Assert.Contains("href=\"#section\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"md-toc", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<nav", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Portable_Html_Fallbacks_Render_Footnotes_Without_OfficeImo_Section_Chrome() {
        var doc = MarkdownReader.Parse("""
Lead[^1]

[^1]: Footnote text
""");
        var options = new HtmlOptions { Kind = HtmlKind.Fragment, BodyClass = null };
        MarkdownBlockRenderBuiltInExtensions.AddPortableHtmlFallbacks(options);

        var html = doc.ToHtmlFragment(options);

        Assert.Contains("<section><hr />", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<p id=\"fn:1\"><sup>1</sup> Footnote text", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"footnotes\"", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<ol>", html, StringComparison.OrdinalIgnoreCase);
    }
}
