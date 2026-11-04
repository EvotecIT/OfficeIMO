using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_PdfFidelity_Parsing_Tests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OrderedListResumesAfterNestedBullets(bool commonMark) {
        const string source = "1. First\n   - Nested one\n   - Nested two\n2. Second\n3. Third";
        var document = MarkdownDocument.Parse(source, commonMark ? MarkdownLoadOptions.CreateCommonMarkProfile() : new MarkdownLoadOptions());
        var list = Assert.IsType<OrderedListBlock>(Assert.Single(document.Blocks));
        Assert.Equal(3, list.Items.Count);
        Assert.Contains("<li>Second</li>", document.ToHtml(new global::OfficeIMO.Markdown.MarkdownHtmlExportOptions { Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document }));
        Assert.Contains("<li>Third</li>", document.ToHtml(new global::OfficeIMO.Markdown.MarkdownHtmlExportOptions { Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document }));
        Assert.DoesNotContain("2. Second", document.ToHtml(new global::OfficeIMO.Markdown.MarkdownHtmlExportOptions { Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document }));
    }

    [Fact]
    public void FootnoteLeadWithColonIsNarrativeAndKeepsStructuredContinuation() {
        var document = MarkdownDocument.Parse("Reference[^1].\n\n[^1]: Data source: dashboard.\n\n  - Detail\n\nTail.");
        var footnote = Assert.Single(document.Blocks.OfType<FootnoteDefinitionBlock>());
        Assert.IsType<ParagraphBlock>(footnote.ChildBlocks[0]);
        Assert.IsType<UnorderedListBlock>(footnote.ChildBlocks[1]);
        Assert.Contains("Data source: dashboard.", document.ToHtml(new global::OfficeIMO.Markdown.MarkdownHtmlExportOptions { Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document }));
        Assert.DoesNotContain("<dt>Data source</dt>", document.ToHtml(new global::OfficeIMO.Markdown.MarkdownHtmlExportOptions { Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Document }));
    }
}
