using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Syntax_Tests {
    [Fact]
    public void ParseWithSyntaxTree_Captures_TopLevel_Block_Kinds_And_Spans() {
        var markdown = """
# Title

Paragraph text
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Equal(MarkdownSyntaxKind.Document, result.SyntaxTree.Kind);
        Assert.Equal(2, result.SyntaxTree.Children.Count);

        var heading = result.SyntaxTree.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.NotNull(heading.SourceSpan);
        Assert.Equal(1, heading.SourceSpan!.Value.StartLine);
        Assert.Equal(1, heading.SourceSpan!.Value.EndLine);
        Assert.Equal("Title", heading.Literal);

        var paragraph = result.SyntaxTree.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.NotNull(paragraph.SourceSpan);
        Assert.Equal(3, paragraph.SourceSpan!.Value.StartLine);
        Assert.Equal(3, paragraph.SourceSpan!.Value.EndLine);
        Assert.Equal("Paragraph text", paragraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Reconstructs_SameType_Nested_Lists() {
        var markdown = """
- parent
  - child
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var list = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);
        Assert.NotNull(list.SourceSpan);
        Assert.Equal(1, list.SourceSpan!.Value.StartLine);
        Assert.Equal(2, list.SourceSpan!.Value.EndLine);

        var parentItem = Assert.Single(list.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, parentItem.Kind);
        Assert.Equal(2, parentItem.Children.Count);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, parentItem.Children[0].Kind);
        Assert.Equal("parent", parentItem.Children[0].Literal);

        var nestedList = parentItem.Children[1];
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, nestedList.Kind);
        var nestedItem = Assert.Single(nestedList.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, nestedItem.Kind);
        var nestedParagraph = Assert.Single(nestedItem.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, nestedParagraph.Kind);
        Assert.Equal("child", nestedParagraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Nested_Quote_Child_Spans() {
        var markdown = """
> quoted
> second
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var quote = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Quote, quote.Kind);
        var paragraph = Assert.Single(quote.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.NotNull(paragraph.SourceSpan);
        Assert.Equal(1, paragraph.SourceSpan!.Value.StartLine);
        Assert.Equal(2, paragraph.SourceSpan!.Value.EndLine);
        Assert.Equal("quoted second", paragraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Nested_Callout_Child_Spans() {
        var markdown = """
> [!NOTE] Title
> body
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var callout = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Callout, callout.Kind);
        Assert.Equal("note:Title", callout.Literal);
        var paragraph = Assert.Single(callout.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.NotNull(paragraph.SourceSpan);
        Assert.Equal(2, paragraph.SourceSpan!.Value.StartLine);
        Assert.Equal(2, paragraph.SourceSpan!.Value.EndLine);
        Assert.Equal("body", paragraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Preserves_Existing_Object_Model_Output() {
        var markdown = """
> quote

Term: Definition
""";

        var expected = MarkdownReader.Parse(markdown);
        var detailed = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Equal(expected.Blocks.Count, detailed.Document.Blocks.Count);
        Assert.Equal(expected.ToMarkdown(), detailed.Document.ToMarkdown());
    }
}
