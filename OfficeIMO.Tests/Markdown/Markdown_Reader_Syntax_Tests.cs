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
        Assert.NotNull(parentItem.SourceSpan);
        Assert.Equal(1, parentItem.SourceSpan!.Value.StartLine);
        Assert.Equal(2, parentItem.SourceSpan!.Value.EndLine);
        Assert.Equal(2, parentItem.Children.Count);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, parentItem.Children[0].Kind);
        Assert.Equal("parent", parentItem.Children[0].Literal);

        var nestedList = parentItem.Children[1];
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, nestedList.Kind);
        Assert.NotNull(nestedList.SourceSpan);
        Assert.Equal(2, nestedList.SourceSpan!.Value.StartLine);
        Assert.Equal(2, nestedList.SourceSpan!.Value.EndLine);
        var nestedItem = Assert.Single(nestedList.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, nestedItem.Kind);
        Assert.NotNull(nestedItem.SourceSpan);
        Assert.Equal(2, nestedItem.SourceSpan!.Value.StartLine);
        Assert.Equal(2, nestedItem.SourceSpan!.Value.EndLine);
        var nestedParagraph = Assert.Single(nestedItem.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, nestedParagraph.Kind);
        Assert.Equal("child", nestedParagraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_ListItem_Child_Spans() {
        var markdown = """
- lead
  continued

  > quoted
  > second

  trailing para
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var list = Assert.Single(result.SyntaxTree.Children);
        var item = Assert.Single(list.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, item.Kind);
        Assert.NotNull(item.SourceSpan);
        Assert.Equal(1, item.SourceSpan!.Value.StartLine);
        Assert.Equal(7, item.SourceSpan!.Value.EndLine);
        Assert.Equal(3, item.Children.Count);

        var leadParagraph = item.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, leadParagraph.Kind);
        Assert.NotNull(leadParagraph.SourceSpan);
        Assert.Equal(1, leadParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(2, leadParagraph.SourceSpan!.Value.EndLine);
        Assert.Equal("lead continued", leadParagraph.Literal);

        var quote = item.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Quote, quote.Kind);
        Assert.NotNull(quote.SourceSpan);
        Assert.Equal(4, quote.SourceSpan!.Value.StartLine);
        Assert.Equal(5, quote.SourceSpan!.Value.EndLine);
        var quoteParagraph = Assert.Single(quote.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, quoteParagraph.Kind);
        Assert.NotNull(quoteParagraph.SourceSpan);
        Assert.Equal(4, quoteParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(5, quoteParagraph.SourceSpan!.Value.EndLine);

        var trailingParagraph = item.Children[2];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, trailingParagraph.Kind);
        Assert.NotNull(trailingParagraph.SourceSpan);
        Assert.Equal(7, trailingParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(7, trailingParagraph.SourceSpan!.Value.EndLine);
        Assert.Equal("trailing para", trailingParagraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Setext_Headings_Inside_List_Items() {
        var markdown = """
- Item title
  ----------

  body
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var list = Assert.Single(result.SyntaxTree.Children);
        var item = Assert.Single(list.Children);
        Assert.Equal(2, item.Children.Count);

        var heading = item.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.NotNull(heading.SourceSpan);
        Assert.Equal(1, heading.SourceSpan!.Value.StartLine);
        Assert.Equal(2, heading.SourceSpan!.Value.EndLine);
        Assert.Equal("Item title", heading.Literal);

        var paragraph = item.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.NotNull(paragraph.SourceSpan);
        Assert.Equal(4, paragraph.SourceSpan!.Value.StartLine);
        Assert.Equal(4, paragraph.SourceSpan!.Value.EndLine);
        Assert.Equal("body", paragraph.Literal);
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
    public void ParseWithSyntaxTree_Captures_ListItem_Spans_Inside_Quotes() {
        var markdown = """
> intro
>
> - item
>   continued
>
>   trailing
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var quote = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Quote, quote.Kind);
        Assert.Equal(2, quote.Children.Count);

        var list = Assert.IsType<MarkdownSyntaxNode>(quote.Children[1]);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);
        Assert.NotNull(list.SourceSpan);
        Assert.Equal(3, list.SourceSpan!.Value.StartLine);
        Assert.Equal(6, list.SourceSpan!.Value.EndLine);

        var item = Assert.Single(list.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, item.Kind);
        Assert.NotNull(item.SourceSpan);
        Assert.Equal(3, item.SourceSpan!.Value.StartLine);
        Assert.Equal(6, item.SourceSpan!.Value.EndLine);

        var lead = item.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, lead.Kind);
        Assert.NotNull(lead.SourceSpan);
        Assert.Equal(3, lead.SourceSpan!.Value.StartLine);
        Assert.Equal(4, lead.SourceSpan!.Value.EndLine);

        var trailing = item.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, trailing.Kind);
        Assert.NotNull(trailing.SourceSpan);
        Assert.Equal(6, trailing.SourceSpan!.Value.StartLine);
        Assert.Equal(6, trailing.SourceSpan!.Value.EndLine);
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
    public void ParseWithSyntaxTree_Captures_ListItem_Spans_Inside_Callouts() {
        var markdown = """
> [!TIP] Title
> - item
>   continued
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var callout = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Callout, callout.Kind);
        var list = Assert.Single(callout.Children);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);
        Assert.NotNull(list.SourceSpan);
        Assert.Equal(2, list.SourceSpan!.Value.StartLine);
        Assert.Equal(3, list.SourceSpan!.Value.EndLine);

        var item = Assert.Single(list.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, item.Kind);
        Assert.NotNull(item.SourceSpan);
        Assert.Equal(2, item.SourceSpan!.Value.StartLine);
        Assert.Equal(3, item.SourceSpan!.Value.EndLine);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Definition_List_Item_Spans() {
        var markdown = """
Term: Definition
Other: Another
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var definitionList = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.DefinitionList, definitionList.Kind);
        Assert.NotNull(definitionList.SourceSpan);
        Assert.Equal(1, definitionList.SourceSpan!.Value.StartLine);
        Assert.Equal(2, definitionList.SourceSpan!.Value.EndLine);

        Assert.Equal(2, definitionList.Children.Count);

        var firstItem = definitionList.Children[0];
        Assert.Equal(MarkdownSyntaxKind.DefinitionItem, firstItem.Kind);
        Assert.NotNull(firstItem.SourceSpan);
        Assert.Equal(1, firstItem.SourceSpan!.Value.StartLine);
        Assert.Equal(1, firstItem.SourceSpan!.Value.EndLine);
        Assert.Equal("Term", firstItem.Literal);

        var firstDefinition = Assert.Single(firstItem.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, firstDefinition.Kind);
        Assert.NotNull(firstDefinition.SourceSpan);
        Assert.Equal(1, firstDefinition.SourceSpan!.Value.StartLine);
        Assert.Equal(1, firstDefinition.SourceSpan!.Value.EndLine);
        Assert.Equal("Definition", firstDefinition.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Details_Body_Child_Spans() {
        var markdown = """
<details>
<summary>Summary</summary>

- item
  continued

</details>
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var details = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Details, details.Kind);
        Assert.Equal(2, details.Children.Count);

        var summary = details.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Summary, summary.Kind);
        Assert.NotNull(summary.SourceSpan);
        Assert.Equal(2, summary.SourceSpan!.Value.StartLine);
        Assert.Equal(2, summary.SourceSpan!.Value.EndLine);

        var list = details.Children[1];
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);
        Assert.NotNull(list.SourceSpan);
        Assert.Equal(4, list.SourceSpan!.Value.StartLine);
        Assert.Equal(5, list.SourceSpan!.Value.EndLine);

        var item = Assert.Single(list.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, item.Kind);
        Assert.NotNull(item.SourceSpan);
        Assert.Equal(4, item.SourceSpan!.Value.StartLine);
        Assert.Equal(5, item.SourceSpan!.Value.EndLine);
    }

    [Fact]
    public void ParseWithSyntaxTree_Finds_Deepest_Node_By_Line() {
        var markdown = """
# Title

- lead
  continued

  > quoted
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var titleNode = result.SyntaxTree.FindDeepestNodeAtLine(1);
        Assert.NotNull(titleNode);
        Assert.Equal(MarkdownSyntaxKind.Heading, titleNode!.Kind);
        Assert.Equal("Title", titleNode.Literal);

        var leadNode = result.SyntaxTree.FindDeepestNodeAtLine(3);
        Assert.NotNull(leadNode);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, leadNode!.Kind);
        Assert.Equal("lead continued", leadNode.Literal);

        var quoteNode = result.SyntaxTree.FindDeepestNodeAtLine(6);
        Assert.NotNull(quoteNode);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, quoteNode!.Kind);
        Assert.Equal("quoted", quoteNode.Literal);

        Assert.Null(result.SyntaxTree.FindDeepestNodeAtLine(99));
    }

    [Fact]
    public void ParseWithSyntaxTree_Enumerates_Descendants_And_Self() {
        var markdown = """
Paragraph
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var kinds = result.SyntaxTree.DescendantsAndSelf().Select(node => node.Kind).ToArray();

        Assert.Equal(new[] { MarkdownSyntaxKind.Document, MarkdownSyntaxKind.Paragraph }, kinds);
    }

    [Fact]
    public void ParseWithSyntaxTree_Finds_Node_Path_By_Line() {
        var markdown = """
> [!TIP] Title
> - item
>   continued
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var path = result.SyntaxTree.FindNodePathAtLine(3).Select(node => node.Kind).ToArray();

        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.Callout,
            MarkdownSyntaxKind.UnorderedList,
            MarkdownSyntaxKind.ListItem,
            MarkdownSyntaxKind.Paragraph
        }, path);

        Assert.Empty(result.SyntaxTree.FindNodePathAtLine(99));
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
