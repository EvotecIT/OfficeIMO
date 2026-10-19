using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_PdfFidelity_Parsing_Tests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OrderedListResumesAfterNestedBullets(bool commonMark) {
        const string source = "1. First\n   - Nested one\n   - Nested two\n2. Second\n3. Third";
        var document = MarkdownReader.Parse(source, commonMark ? MarkdownReaderOptions.CreateCommonMarkProfile() : new MarkdownReaderOptions());
        var list = Assert.IsType<OrderedListBlock>(Assert.Single(document.Blocks));
        Assert.Equal(3, list.Items.Count);
        Assert.Contains("<li>Second</li>", document.ToHtmlDocument());
        Assert.Contains("<li>Third</li>", document.ToHtmlDocument());
        Assert.DoesNotContain("2. Second", document.ToHtmlDocument());
    }

    [Fact]
    public void FootnoteLeadWithColonIsNarrativeAndKeepsStructuredContinuation() {
        var document = MarkdownReader.Parse("Reference[^1].\n\n[^1]: Data source: dashboard.\n\n  - Detail\n\nTail.");
        var footnote = Assert.Single(document.Blocks.OfType<FootnoteDefinitionBlock>());
        Assert.IsType<ParagraphBlock>(footnote.ChildBlocks[0]);
        Assert.IsType<UnorderedListBlock>(footnote.ChildBlocks[1]);
        Assert.Contains("Data source: dashboard.", document.ToHtmlDocument());
        Assert.DoesNotContain("<dt>Data source</dt>", document.ToHtmlDocument());
    }
}
