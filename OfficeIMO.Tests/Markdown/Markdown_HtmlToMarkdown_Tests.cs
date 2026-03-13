using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownHtmlToMarkdownTests {
    [Fact]
    public void HtmlToMarkdown_Converter_UsesDefaultOptionsWhenNull() {
        var converter = new HtmlToMarkdownConverter();

        string markdown = converter.Convert("<p>Hello</p>", options: null);
        MarkdownDoc document = converter.ConvertToDocument("<p>Hello</p>", options: null);

        Assert.Contains("Hello", markdown, StringComparison.Ordinal);
        Assert.Single(document.Blocks);
        Assert.IsType<ParagraphBlock>(document.Blocks[0]);
    }

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
    public void HtmlToMarkdown_ConvertsHtmlFragmentWithoutBodyWrapper() {
        string html = "<h2>Fragment</h2><p>Body</p>";

        MarkdownDoc document = html.LoadFromHtml();

        Assert.Collection(document.Blocks,
            block => Assert.IsType<HeadingBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
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

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedInlineHtml_WhenRequested() {
        string html = "<p>Hello <custom-inline data-name=\"demo\">world</custom-inline></p>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            PreserveUnsupportedInlineHtml = true
        });

        Assert.Contains("<custom-inline", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_CapturesFigureCaptionOnImageBlocks() {
        string html = "<figure><img src=\"/img/demo.png\" alt=\"Demo\" /><figcaption>Example caption</figcaption></figure>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("https://example.com/img/demo.png", image.Path);
        Assert.Equal("Demo", image.Alt);
        Assert.Equal("Example caption", image.Caption);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMixedListItemBlockOrder() {
        string html = "<ul><li><p>Alpha</p><blockquote><p>Quoted</p></blockquote><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<QuoteBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));

        string markdown = document.ToMarkdown();
        int alphaIndex = markdown.IndexOf("Alpha", StringComparison.Ordinal);
        int quoteIndex = markdown.IndexOf("Quoted", StringComparison.Ordinal);
        int omegaIndex = markdown.IndexOf("Omega", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha in markdown output.");
        Assert.True(quoteIndex > alphaIndex, "Expected quoted content after the opening paragraph.");
        Assert.True(omegaIndex > quoteIndex, "Expected trailing paragraph after the quote block.");
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMultipleDefinitionsPerTerm() {
        string html = "<dl><dt>Term</dt><dd>First definition</dd><dd>Second definition</dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Equal(2, list.Items.Count);
        Assert.Equal("Term", list.Items[0].Term);
        Assert.Equal("First definition", list.Items[0].Definition);
        Assert.Equal("Term", list.Items[1].Term);
        Assert.Equal("Second definition", list.Items[1].Definition);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesNestedListOrderInsideListItem() {
        string html = "<ul><li><p>Alpha</p><ul><li>Nested</li></ul><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<UnorderedListBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));

        string markdown = document.ToMarkdown();
        int alphaIndex = markdown.IndexOf("Alpha", StringComparison.Ordinal);
        int nestedIndex = markdown.IndexOf("Nested", StringComparison.Ordinal);
        int omegaIndex = markdown.IndexOf("Omega", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha in markdown output.");
        Assert.True(nestedIndex > alphaIndex, "Expected nested list content after the opening paragraph.");
        Assert.True(omegaIndex > nestedIndex, "Expected trailing paragraph after the nested list.");
    }

    [Fact]
    public void HtmlToMarkdown_PreservesDetailsOrderInsideListItem() {
        string html = "<ul><li><p>Alpha</p><details open><summary>More</summary><p>Hidden</p></details><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<DetailsBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMultipleTermsPerDefinitionGroup() {
        string html = "<dl><dt>Alpha</dt><dt>Beta</dt><dd>Shared definition</dd><dd>Follow-up definition</dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Equal(4, list.Items.Count);
        Assert.Equal(("Alpha", "Shared definition"), list.Items[0]);
        Assert.Equal(("Beta", "Shared definition"), list.Items[1]);
        Assert.Equal(("Alpha", "Follow-up definition"), list.Items[2]);
        Assert.Equal(("Beta", "Follow-up definition"), list.Items[3]);
    }
}
