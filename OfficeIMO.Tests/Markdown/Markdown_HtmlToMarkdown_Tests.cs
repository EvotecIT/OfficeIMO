using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownHtmlToMarkdownTests {
    [Fact]
    public void HtmlToMarkdown_ConvertsCommonDocumentBlocks() {
        string html = "<html><body><h1>Hello</h1><p>A <strong>bold</strong> <a href=\"https://example.com\">link</a>.</p><ul><li>One</li><li>Two</li></ul></body></html>";

        string markdown = html.ToMarkdown();

        Assert.Contains("# Hello", markdown, StringComparison.Ordinal);
        Assert.Contains("**bold**", markdown, StringComparison.Ordinal);
        Assert.Contains("[link](https://example.com)", markdown, StringComparison.Ordinal);
        Assert.Contains("- One", markdown, StringComparison.Ordinal);
        Assert.Contains("- Two", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_LoadFromHtml_ProducesTypedBlocks() {
        string html = "<html><body><h2>Section</h2><blockquote><p>Quoted</p></blockquote><details open><summary>More</summary><p>Hidden text</p></details></body></html>";

        MarkdownDoc document = html.LoadFromHtml();

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Level == 2 && heading.Text == "Section");
        Assert.Contains(document.Blocks, block => block is QuoteBlock);
        Assert.Contains(document.Blocks, block => block is DetailsBlock details && details.Open);
    }

    [Fact]
    public void HtmlToMarkdown_ResolvesRelativeLinksWithBaseUri() {
        string html = "<p><a href=\"guide/start\">Docs</a></p>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/docs/")
        });

        Assert.Contains("[Docs](https://example.com/docs/guide/start)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlocks_WhenRequested() {
        string html = "<custom-widget data-name=\"demo\">Hello</custom-widget>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        Assert.Contains("<custom-widget", markdown, StringComparison.Ordinal);
    }
}
