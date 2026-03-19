using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
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
        Assert.NotNull(result.SyntaxTree.SourceSpan);
        Assert.Equal(1, result.SyntaxTree.SourceSpan!.Value.StartLine);
        Assert.Equal(3, result.SyntaxTree.SourceSpan!.Value.EndLine);
        Assert.Equal(2, result.SyntaxTree.Children.Count);

        var heading = result.SyntaxTree.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.NotNull(heading.SourceSpan);
        Assert.Equal(1, heading.SourceSpan!.Value.StartLine);
        Assert.Equal(1, heading.SourceSpan!.Value.EndLine);
        Assert.Equal(1, heading.SourceSpan!.Value.StartColumn);
        Assert.Equal(7, heading.SourceSpan!.Value.EndColumn);
        Assert.Equal("Title", heading.Literal);

        var paragraph = result.SyntaxTree.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.NotNull(paragraph.SourceSpan);
        Assert.Equal(3, paragraph.SourceSpan!.Value.StartLine);
        Assert.Equal(3, paragraph.SourceSpan!.Value.EndLine);
        Assert.Equal(1, paragraph.SourceSpan!.Value.StartColumn);
        Assert.Equal(14, paragraph.SourceSpan!.Value.EndColumn);
        Assert.Equal("Paragraph text", paragraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Handles_Mixed_Line_Endings_Without_Trailing_Newline() {
        const string markdown = "# Title\r\n\r\nParagraph one\r\rSecond para";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Equal(3, result.SyntaxTree.Children.Count);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 5, 11), result.SyntaxTree.SourceSpan);

        var heading = result.SyntaxTree.Children[0];
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 7), heading.SourceSpan);

        var firstParagraph = result.SyntaxTree.Children[1];
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 13), firstParagraph.SourceSpan);

        var secondParagraph = result.SyntaxTree.Children[2];
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 11), secondParagraph.SourceSpan);
        Assert.Equal("Second para", secondParagraph.Literal);
        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(5, 7)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Returns_FinalDocument_OriginalSyntaxTree_And_TransformDiagnostics() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());
        const string markdown = "previous shutdown was unexpected### Reason";

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

        Assert.Equal(2, result.Document.Blocks.Count);
        Assert.Single(result.SyntaxTree.Children);
        var diagnostic = Assert.Single(result.TransformDiagnostics);
        Assert.Contains(nameof(MarkdownCompactHeadingBoundaryTransform), diagnostic.TransformName, StringComparison.Ordinal);
        Assert.Equal(0, diagnostic.ChangedBlockStartBefore);
        Assert.Equal(1, diagnostic.ChangedBlockCountBefore);
        Assert.Equal(0, diagnostic.ChangedBlockStartAfter);
        Assert.Equal(2, diagnostic.ChangedBlockCountAfter);
        Assert.Equal(new MarkdownSourceSpan(1, 1), diagnostic.AffectedSourceSpan);
        Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(2, result.FinalSyntaxTree.Children.Count);
        Assert.Equal(MarkdownSyntaxKind.Heading, result.FinalSyntaxTree.Children[1].Kind);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Provides_Final_Syntax_Lookup_Helpers() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteFirstParagraphTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("hello", options);

        Assert.Equal("hello", result.FindDeepestNodeAtLine(1)!.Literal);
        Assert.Equal("rewritten", result.FindDeepestFinalNodeAtLine(1)!.Literal);
        Assert.Equal("hello", result.FindDeepestNodeContainingSpan(new MarkdownSourceSpan(1, 1))!.Literal);
        Assert.Equal("rewritten", result.FindDeepestFinalNodeContainingSpan(new MarkdownSourceSpan(1, 1))!.Literal);
        Assert.Equal(new[] { MarkdownSyntaxKind.Document, MarkdownSyntaxKind.Paragraph }, result.FindFinalNodePathAtLine(1).Select(node => node.Kind).ToArray());
        Assert.Equal("rewritten", result.FindNearestFinalBlockOverlappingSpan(new MarkdownSourceSpan(1, 1))!.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Heading_Structure() {
        var markdown = """
Heading Title
-------------
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var heading = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.Equal("Heading Title", heading.Literal);

        var level = heading.Children[0];
        Assert.Equal(MarkdownSyntaxKind.HeadingLevel, level.Kind);
        Assert.Equal("2", level.Literal);
        Assert.Null(level.SourceSpan);

        var text = heading.Children[1];
        Assert.Equal(MarkdownSyntaxKind.HeadingText, text.Kind);
        Assert.Equal("Heading Title", text.Literal);
        Assert.NotNull(text.SourceSpan);
        Assert.Equal(1, text.SourceSpan!.Value.StartLine);
        Assert.Equal(1, text.SourceSpan!.Value.EndLine);
    }

    [Fact]
    public void ParseWithSyntaxTree_Preserves_Heading_Inline_Markup_In_Literals() {
        const string markdown = "# **Heading** `Text`";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var heading = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.Equal("**Heading** `Text`", heading.Literal);

        var text = heading.Children[1];
        Assert.Equal(MarkdownSyntaxKind.HeadingText, text.Kind);
        Assert.Equal("**Heading** `Text`", text.Literal);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineStrong,
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineCodeSpan
        }, text.Children.Select(node => node.Kind).ToArray());
        Assert.Equal(5, text.SourceSpan!.Value.StartColumn);
        Assert.Equal(20, text.SourceSpan!.Value.EndColumn);
        Assert.NotNull(text.Children[0].SourceSpan);
        Assert.Equal(5, text.Children[0].SourceSpan!.Value.StartColumn);
        Assert.Equal(11, text.Children[0].SourceSpan!.Value.EndColumn);
        Assert.NotNull(text.Children[1].SourceSpan);
        Assert.Equal(14, text.Children[1].SourceSpan!.Value.StartColumn);
        Assert.Equal(14, text.Children[1].SourceSpan!.Value.EndColumn);
        Assert.NotNull(text.Children[2].SourceSpan);
        Assert.Equal(15, text.Children[2].SourceSpan!.Value.StartColumn);
        Assert.Equal(20, text.Children[2].SourceSpan!.Value.EndColumn);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Paragraph_Inline_Syntax_Structure() {
        const string markdown = "Use **bold** [docs](https://example.com) and `code`.";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var paragraph = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Equal("Use **bold** [docs](https://example.com) and `code`.", paragraph.Literal);

        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineStrong,
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineLink,
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineCodeSpan,
            MarkdownSyntaxKind.InlineText
        }, paragraph.Children.Select(node => node.Kind).ToArray());

        var strong = paragraph.Children[1];
        Assert.Equal("bold", strong.Literal);

        var link = paragraph.Children[3];
        Assert.Equal("https://example.com", link.Literal);
        Assert.Single(link.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineText, link.Children[0].Kind);
        Assert.Equal("docs", link.Children[0].Literal);

        var code = paragraph.Children[5];
        Assert.Equal(MarkdownSyntaxKind.InlineCodeSpan, code.Kind);
        Assert.Equal("code", code.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Inline_SourceSpans_And_Position_Lookups() {
        const string markdown = "Use **bold** [docs](https://example.com) and `code`.";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var paragraph = Assert.Single(result.SyntaxTree.Children);

        Assert.Equal(1, paragraph.SourceSpan!.Value.StartColumn);
        Assert.Equal(markdown.Length, paragraph.SourceSpan!.Value.EndColumn);

        var strong = paragraph.Children[1];
        Assert.Equal(7, strong.SourceSpan!.Value.StartColumn);
        Assert.Equal(10, strong.SourceSpan!.Value.EndColumn);

        var link = paragraph.Children[3];
        Assert.Equal(14, link.SourceSpan!.Value.StartColumn);
        Assert.Equal(40, link.SourceSpan!.Value.EndColumn);

        var code = paragraph.Children[5];
        Assert.Equal(46, code.SourceSpan!.Value.StartColumn);
        Assert.Equal(51, code.SourceSpan!.Value.EndColumn);

        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(1, 8)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.InlineLink, result.FindDeepestNodeAtPosition(1, 30)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.InlineCodeSpan, result.FindDeepestNodeAtPosition(1, 48)!.Kind);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineLink
        }, result.FindNodePathAtPosition(1, 30).Select(node => node.Kind).ToArray());
        Assert.Equal(MarkdownSyntaxKind.Paragraph, result.FindNearestBlockAtPosition(1, 48)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Paragraph_Image_And_HardBreak_Inline_Nodes() {
        const string markdown = "See ![Alt](image.png \"Title\")  \nnext";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var paragraph = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineImage,
            MarkdownSyntaxKind.InlineHardBreak,
            MarkdownSyntaxKind.InlineText
        }, paragraph.Children.Select(node => node.Kind).ToArray());

        var image = paragraph.Children[1];
        Assert.Equal("image.png", image.Literal);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.ImageAlt,
            MarkdownSyntaxKind.ImageSource,
            MarkdownSyntaxKind.ImageTitle
        }, image.Children.Select(node => node.Kind).ToArray());
        Assert.Equal("Alt", image.Children[0].Literal);
        Assert.Equal("image.png", image.Children[1].Literal);
        Assert.Equal("Title", image.Children[2].Literal);
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
        var leadText = Assert.Single(leadParagraph.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineText, leadText.Kind);
        Assert.NotNull(leadText.SourceSpan);
        Assert.Equal(1, leadText.SourceSpan!.Value.StartLine);
        Assert.Equal(3, leadText.SourceSpan!.Value.StartColumn);
        Assert.Equal(2, leadText.SourceSpan!.Value.EndLine);
        Assert.Equal(11, leadText.SourceSpan!.Value.EndColumn);

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
        var trailingText = Assert.Single(trailingParagraph.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineText, trailingText.Kind);
        Assert.NotNull(trailingText.SourceSpan);
        Assert.Equal(7, trailingText.SourceSpan!.Value.StartLine);
        Assert.Equal(3, trailingText.SourceSpan!.Value.StartColumn);
        Assert.Equal(7, trailingText.SourceSpan!.Value.EndLine);
        Assert.Equal(15, trailingText.SourceSpan!.Value.EndColumn);

        var deepLead = result.FindDeepestNodeAtPosition(2, 4);
        Assert.NotNull(deepLead);
        Assert.Equal(MarkdownSyntaxKind.InlineText, deepLead!.Kind);
        Assert.Equal("lead continued", deepLead.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Loose_List_Item_Trailing_Paragraph_SourceSpans() {
        var markdown = """
- item
  continued

  trailing
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var list = Assert.Single(result.SyntaxTree.Children);
        var item = Assert.Single(list.Children);
        Assert.Equal(2, item.Children.Count);

        var lead = item.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, lead.Kind);
        Assert.Equal(1, lead.SourceSpan!.Value.StartLine);
        Assert.Equal(3, lead.SourceSpan!.Value.StartColumn);
        Assert.Equal(2, lead.SourceSpan!.Value.EndLine);
        Assert.Equal(11, lead.SourceSpan!.Value.EndColumn);

        var trailing = item.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, trailing.Kind);
        Assert.Equal(4, trailing.SourceSpan!.Value.StartLine);
        Assert.Equal(3, trailing.SourceSpan!.Value.StartColumn);
        Assert.Equal(4, trailing.SourceSpan!.Value.EndLine);
        Assert.Equal(10, trailing.SourceSpan!.Value.EndColumn);
        var trailingText = Assert.Single(trailing.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineText, trailingText.Kind);
        Assert.Equal(3, trailingText.SourceSpan!.Value.StartColumn);
        Assert.Equal(10, trailingText.SourceSpan!.Value.EndColumn);
        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(4, 4)!.Kind);
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
    public void ParseWithSyntaxTree_Captures_Trailing_Paragraph_After_List_Item_Setext_Heading() {
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
        Assert.Equal(3, paragraph.SourceSpan!.Value.StartLine);
        Assert.Equal(3, paragraph.SourceSpan!.Value.EndLine);
        Assert.Equal("body", paragraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Separates_Blank_Line_Before_List_Item_Setext_Heading() {
        var markdown = """
- Item

  Heading
  -------
  body
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var list = Assert.Single(result.SyntaxTree.Children);
        var item = Assert.Single(list.Children);
        Assert.Equal(3, item.Children.Count);

        var firstParagraph = item.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, firstParagraph.Kind);
        Assert.NotNull(firstParagraph.SourceSpan);
        Assert.Equal(1, firstParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(1, firstParagraph.SourceSpan!.Value.EndLine);
        Assert.Equal("Item", firstParagraph.Literal);

        var heading = item.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.NotNull(heading.SourceSpan);
        Assert.Equal(3, heading.SourceSpan!.Value.StartLine);
        Assert.Equal(4, heading.SourceSpan!.Value.EndLine);
        Assert.Equal("Heading", heading.Literal);

        var trailingParagraph = item.Children[2];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, trailingParagraph.Kind);
        Assert.NotNull(trailingParagraph.SourceSpan);
        Assert.Equal(5, trailingParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(5, trailingParagraph.SourceSpan!.Value.EndLine);
        Assert.Equal("body", trailingParagraph.Literal);
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
        Assert.Equal(3, paragraph.SourceSpan!.Value.StartColumn);
        Assert.Equal(8, paragraph.SourceSpan!.Value.EndColumn);
        Assert.Equal("quoted second", paragraph.Literal);
        var text = Assert.Single(paragraph.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineText, text.Kind);
        Assert.NotNull(text.SourceSpan);
        Assert.Equal(1, text.SourceSpan!.Value.StartLine);
        Assert.Equal(3, text.SourceSpan!.Value.StartColumn);
        Assert.Equal(2, text.SourceSpan!.Value.EndLine);
        Assert.Equal(8, text.SourceSpan!.Value.EndColumn);
        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(2, 4)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Rebuilds_Final_Quote_Syntax_After_Nested_Transform() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
> original
> second
""", options);

        Assert.Equal("original second", result.FindDeepestNodeAtPosition(1, 4)!.Literal);

        var finalQuote = Assert.Single(result.FinalSyntaxTree.Children);
        var finalParagraph = Assert.Single(finalQuote.Children);
        var finalText = Assert.Single(finalParagraph.Children);

        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
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
        Assert.Equal(5, lead.SourceSpan!.Value.StartColumn);
        Assert.Equal(13, lead.SourceSpan!.Value.EndColumn);
        var leadText = Assert.Single(lead.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineText, leadText.Kind);
        Assert.Equal(5, leadText.SourceSpan!.Value.StartColumn);
        Assert.Equal(13, leadText.SourceSpan!.Value.EndColumn);

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
        Assert.Equal(3, paragraph.SourceSpan!.Value.StartColumn);
        Assert.Equal(6, paragraph.SourceSpan!.Value.EndColumn);
        Assert.Equal("body", paragraph.Literal);
        var text = Assert.Single(paragraph.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineText, text.Kind);
        Assert.Equal(3, text.SourceSpan!.Value.StartColumn);
        Assert.Equal(6, text.SourceSpan!.Value.EndColumn);
        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(2, 4)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Rebuilds_Final_Callout_Syntax_After_Nested_Transform() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
> [!NOTE] Title
> original
""", options);

        Assert.Equal("original", result.FindDeepestNodeAtPosition(2, 4)!.Literal);

        var finalCallout = Assert.Single(result.FinalSyntaxTree.Children);
        var finalParagraph = Assert.Single(finalCallout.Children);
        var finalText = Assert.Single(finalParagraph.Children);

        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
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
        var lead = Assert.Single(item.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, lead.Kind);
        Assert.Equal(5, lead.SourceSpan!.Value.StartColumn);
        Assert.Equal(13, lead.SourceSpan!.Value.EndColumn);
    }

    [Fact]
    public void ParseWithSyntaxTree_Preserves_Callout_Title_Inline_Markup_In_Literal() {
        var markdown = """
> [!NOTE] Title with **strong** [link](https://example.com)
> body
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var callout = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Callout, callout.Kind);
        Assert.Equal("note:Title with **strong** [link](https://example.com)", callout.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Definition_List_Group_Spans() {
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

        var firstGroup = definitionList.Children[0];
        Assert.Equal(MarkdownSyntaxKind.DefinitionGroup, firstGroup.Kind);
        Assert.NotNull(firstGroup.SourceSpan);
        Assert.Equal(1, firstGroup.SourceSpan!.Value.StartLine);
        Assert.Equal(1, firstGroup.SourceSpan!.Value.EndLine);
        Assert.Null(firstGroup.Literal);
        Assert.Equal(2, firstGroup.Children.Count);

        var firstTerm = firstGroup.Children[0];
        Assert.Equal(MarkdownSyntaxKind.DefinitionTerm, firstTerm.Kind);
        Assert.NotNull(firstTerm.SourceSpan);
        Assert.Equal(1, firstTerm.SourceSpan!.Value.StartLine);
        Assert.Equal(1, firstTerm.SourceSpan!.Value.EndLine);
        Assert.Equal("Term", firstTerm.Literal);

        var firstValue = firstGroup.Children[1];
        Assert.Equal(MarkdownSyntaxKind.DefinitionValue, firstValue.Kind);
        Assert.NotNull(firstValue.SourceSpan);
        Assert.Equal(1, firstValue.SourceSpan!.Value.StartLine);
        Assert.Equal(1, firstValue.SourceSpan!.Value.EndLine);
        Assert.Equal("Definition", firstValue.Literal);

        var firstDefinition = Assert.Single(firstValue.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, firstDefinition.Kind);
        Assert.NotNull(firstDefinition.SourceSpan);
        Assert.Equal(1, firstDefinition.SourceSpan!.Value.StartLine);
        Assert.Equal(1, firstDefinition.SourceSpan!.Value.EndLine);
        Assert.Equal("Definition", firstDefinition.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Definition_List_Inline_Structure_And_Position_Lookups() {
        var markdown = """
**Term**: Use [docs](https://example.com)
Other: `code`
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var definitionList = Assert.Single(result.SyntaxTree.Children);
        var firstGroup = definitionList.Children[0];
        var firstTerm = firstGroup.Children[0];
        var firstValue = firstGroup.Children[1];
        var firstParagraph = Assert.Single(firstValue.Children);

        Assert.Equal(new[] { MarkdownSyntaxKind.InlineStrong }, firstTerm.Children.Select(node => node.Kind).ToArray());
        Assert.Equal(1, firstTerm.SourceSpan!.Value.StartColumn);
        Assert.Equal(8, firstTerm.SourceSpan!.Value.EndColumn);
        Assert.Equal(3, firstTerm.Children[0].SourceSpan!.Value.StartColumn);
        Assert.Equal(6, firstTerm.Children[0].SourceSpan!.Value.EndColumn);

        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineLink
        }, firstParagraph.Children.Select(node => node.Kind).ToArray());
        Assert.Equal(11, firstValue.SourceSpan!.Value.StartColumn);
        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(1, 4)!.Kind);
        Assert.Equal("https://example.com", result.FindDeepestNodeAtPosition(1, 20)!.Literal);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.DefinitionList,
            MarkdownSyntaxKind.DefinitionGroup,
            MarkdownSyntaxKind.DefinitionValue,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineLink
        }, result.FindNodePathAtPosition(1, 20).Select(node => node.Kind).ToArray());
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Rebuilds_Final_Definition_List_Syntax_After_Transform() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteDefinitionListDefinitionsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
Term: original
Other: second
""", options);

        Assert.Equal("original", result.FindDeepestNodeAtPosition(1, 7)!.Literal);

        var finalDefinitionList = Assert.Single(result.FinalSyntaxTree.Children);
        var finalFirstGroup = finalDefinitionList.Children[0];
        var finalValue = finalFirstGroup.Children[1];
        var finalParagraph = Assert.Single(finalValue.Children);
        var finalText = Assert.Single(finalParagraph.Children);

        Assert.Equal("rewritten", finalValue.Literal);
        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Multiline_Definition_List_Body_Spans_And_Nested_Blocks() {
        var markdown = """
Term: Intro

  - first
  - second
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var definitionList = Assert.Single(result.SyntaxTree.Children);
        var group = Assert.Single(definitionList.Children);
        var value = group.Children[1];

        Assert.Equal(2, value.Children.Count);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, value.Children[0].Kind);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 11), value.Children[0].SourceSpan);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, value.Children[1].Kind);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 4, 10), value.Children[1].SourceSpan);

        Assert.Equal("first", result.FindDeepestNodeAtPosition(3, 5)!.Literal);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.DefinitionList,
            MarkdownSyntaxKind.DefinitionGroup,
            MarkdownSyntaxKind.DefinitionValue,
            MarkdownSyntaxKind.UnorderedList,
            MarkdownSyntaxKind.ListItem,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineText
        }, result.FindNodePathAtPosition(3, 5).Select(node => node.Kind).ToArray());
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
    public void ParseWithSyntaxTreeAndDiagnostics_Rebuilds_Final_Details_Syntax_After_Nested_Transform() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
<details>
<summary>Summary</summary>

original
</details>
""", options);

        Assert.Equal("original", result.FindDeepestNodeAtPosition(4, 2)!.Literal);

        var finalDetails = Assert.Single(result.FinalSyntaxTree.Children);
        Assert.Equal(2, finalDetails.Children.Count);
        var finalParagraph = finalDetails.Children[1];
        var finalText = Assert.Single(finalParagraph.Children);

        Assert.Equal(MarkdownSyntaxKind.Paragraph, finalParagraph.Kind);
        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Footnote_Paragraph_Spans() {
        var markdown = """
Lead[^1]

[^1]: first line
  continued

  second paragraph
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var footnote = Assert.Single(result.SyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.FootnoteDefinition);
        Assert.NotNull(footnote.SourceSpan);
        Assert.Equal(3, footnote.SourceSpan!.Value.StartLine);
        Assert.Equal(6, footnote.SourceSpan!.Value.EndLine);
        Assert.Equal("1", footnote.Literal);
        Assert.Equal(2, footnote.Children.Count);

        var firstParagraph = footnote.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, firstParagraph.Kind);
        Assert.NotNull(firstParagraph.SourceSpan);
        Assert.Equal(3, firstParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(4, firstParagraph.SourceSpan!.Value.EndLine);
        Assert.Equal("first line continued", firstParagraph.Literal);

        var secondParagraph = footnote.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, secondParagraph.Kind);
        Assert.NotNull(secondParagraph.SourceSpan);
        Assert.Equal(6, secondParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(6, secondParagraph.SourceSpan!.Value.EndLine);
        Assert.Equal("second paragraph", secondParagraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Footnote_Nested_Block_Spans() {
        var markdown = """
Lead[^1]

[^1]: Intro

  - first
  - second
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var footnote = Assert.Single(result.SyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.FootnoteDefinition);
        Assert.NotNull(footnote.SourceSpan);
        Assert.Equal(3, footnote.SourceSpan!.Value.StartLine);
        Assert.Equal(6, footnote.SourceSpan!.Value.EndLine);
        Assert.Equal(2, footnote.Children.Count);

        var intro = footnote.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, intro.Kind);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 11), intro.SourceSpan);

        var list = footnote.Children[1];
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);
        Assert.Equal(new MarkdownSourceSpan(5, 3, 6, 10), list.SourceSpan);

        Assert.Equal("first", result.FindDeepestNodeAtPosition(5, 5)!.Literal);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.FootnoteDefinition,
            MarkdownSyntaxKind.UnorderedList,
            MarkdownSyntaxKind.ListItem,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineText
        }, result.FindNodePathAtPosition(5, 5).Select(node => node.Kind).ToArray());
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Rebuilds_Final_Footnote_Syntax_After_Nested_Transform() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
Lead[^1]

[^1]: original
""", options);

        var finalFootnote = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.FootnoteDefinition);
        var finalParagraph = Assert.Single(finalFootnote.Children);
        var finalText = Assert.Single(finalParagraph.Children);

        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Table_Row_Spans() {
        var markdown = """
| Name | Value |
| --- | ---: |
| One | 1 |
| Two | 2 |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var table = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Table, table.Kind);
        Assert.NotNull(table.SourceSpan);
        Assert.Equal(1, table.SourceSpan!.Value.StartLine);
        Assert.Equal(4, table.SourceSpan!.Value.EndLine);
        Assert.Equal(3, table.Children.Count);

        var header = table.Children[0];
        Assert.Equal(MarkdownSyntaxKind.TableHeader, header.Kind);
        Assert.NotNull(header.SourceSpan);
        Assert.Equal(1, header.SourceSpan!.Value.StartLine);
        Assert.Equal(1, header.SourceSpan!.Value.EndLine);
        Assert.Equal("Name | Value", header.Literal);

        var firstRow = table.Children[1];
        Assert.Equal(MarkdownSyntaxKind.TableRow, firstRow.Kind);
        Assert.NotNull(firstRow.SourceSpan);
        Assert.Equal(3, firstRow.SourceSpan!.Value.StartLine);
        Assert.Equal(3, firstRow.SourceSpan!.Value.EndLine);
        Assert.Equal("One | 1", firstRow.Literal);

        var secondRow = table.Children[2];
        Assert.Equal(MarkdownSyntaxKind.TableRow, secondRow.Kind);
        Assert.NotNull(secondRow.SourceSpan);
        Assert.Equal(4, secondRow.SourceSpan!.Value.StartLine);
        Assert.Equal(4, secondRow.SourceSpan!.Value.EndLine);
        Assert.Equal("Two | 2", secondRow.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Table_Cell_Nodes_And_Cell_Block_Content() {
        var markdown = """
| Name | Notes |
| --- | --- |
| One | Intro<br><br>- first<br>- second |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var table = Assert.Single(result.SyntaxTree.Children);
        var header = table.Children[0];
        Assert.Equal(2, header.Children.Count);
        Assert.All(header.Children, cell => Assert.Equal(MarkdownSyntaxKind.TableCell, cell.Kind));
        Assert.Equal("Name", header.Children[0].Literal);
        Assert.Equal("Notes", header.Children[1].Literal);

        var row = table.Children[1];
        Assert.Equal(2, row.Children.Count);
        Assert.All(row.Children, cell => Assert.Equal(MarkdownSyntaxKind.TableCell, cell.Kind));
        Assert.Equal("One", row.Children[0].Literal);
        Assert.Equal("Intro<br><br>- first<br>- second", row.Children[1].Literal);

        var noteBlocks = row.Children[1].Children;
        Assert.Equal(2, noteBlocks.Count);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, noteBlocks[0].Kind);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, noteBlocks[1].Kind);
        Assert.All(noteBlocks, block => Assert.Equal(3, block.SourceSpan!.Value.StartLine));
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Table_Cell_SourceSpans_And_Position_Lookups() {
        var markdown = """
| Name | Notes |
| --- | --- |
| One | Intro<br><br>- first<br>- second |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var table = Assert.Single(result.SyntaxTree.Children);
        var row = table.Children[1];
        var valueCell = row.Children[1];

        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 5), row.Children[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 9, 3, 40), valueCell.SourceSpan);

        var intro = valueCell.Children[0];
        Assert.Equal(new MarkdownSourceSpan(3, 9, 3, 13), intro.SourceSpan);

        var list = valueCell.Children[1];
        Assert.Equal(new MarkdownSourceSpan(3, 22, 3, 40), list.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 24, 3, 28), list.Children[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 35, 3, 40), list.Children[1].SourceSpan);

        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(3, 3)!.Kind);
        Assert.Equal("One", result.FindDeepestNodeAtPosition(3, 3)!.Literal);
        Assert.Equal("Intro", result.FindDeepestNodeAtPosition(3, 10)!.Literal);
        Assert.Equal("first", result.FindDeepestNodeAtPosition(3, 24)!.Literal);
        Assert.Equal("second", result.FindDeepestNodeAtPosition(3, 36)!.Literal);
    }

    [Fact]
    public void Table_Cells_Expose_Row_Column_Metadata_And_Targeted_Accessors() {
        var markdown = """
| Name | Value |
| --- | --- |
| One | 1 |
| Two | 2 |
""";

        var document = MarkdownReader.Parse(markdown);
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        var header = table.GetHeaderCell(1);
        Assert.NotNull(header);
        Assert.True(header!.IsHeader);
        Assert.Equal(-1, header.RowIndex);
        Assert.Equal(1, header.ColumnIndex);

        var body = table.GetCell(1, 0);
        Assert.NotNull(body);
        Assert.False(body!.IsHeader);
        Assert.Equal(1, body.RowIndex);
        Assert.Equal(0, body.ColumnIndex);

        var cells = table.EnumerateCells().ToArray();
        Assert.Equal(6, cells.Length);
        Assert.Equal(new[] { -1, -1, 0, 0, 1, 1 }, cells.Select(cell => cell.RowIndex).ToArray());
        Assert.Equal(new[] { 0, 1, 0, 1, 0, 1 }, cells.Select(cell => cell.ColumnIndex).ToArray());
    }

    [Fact]
    public void Document_Can_Enumerate_Descendant_Tables_And_Table_Cells() {
        var markdown = """
> | Name | Value |
> | --- | --- |
> | One | 1 |
""";

        var document = MarkdownReader.Parse(markdown);

        var table = Assert.Single(document.DescendantTables());
        Assert.Single(document.DescendantsAndSelf().OfType<QuoteBlock>());

        var cells = document.DescendantTableCells().ToArray();
        Assert.Equal(4, cells.Length);
        Assert.True(cells[0].IsHeader);
        Assert.Equal(-1, cells[0].RowIndex);
        Assert.Equal(0, cells[0].ColumnIndex);
        Assert.False(cells[2].IsHeader);
        Assert.Equal(0, cells[2].RowIndex);
        Assert.Equal(0, cells[2].ColumnIndex);
        var targetedCell = table.GetCell(0, 1);
        Assert.NotNull(targetedCell);
        Assert.Equal(cells[3].Markdown, targetedCell!.Markdown);
        Assert.Equal(cells[3].RowIndex, targetedCell.RowIndex);
        Assert.Equal(cells[3].ColumnIndex, targetedCell.ColumnIndex);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Headerless_Table_Row_Spans() {
        var markdown = """
| One | 1 |
| Two | 2 |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var table = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Table, table.Kind);
        Assert.Equal(2, table.Children.Count);

        var firstRow = table.Children[0];
        Assert.Equal(MarkdownSyntaxKind.TableRow, firstRow.Kind);
        Assert.NotNull(firstRow.SourceSpan);
        Assert.Equal(1, firstRow.SourceSpan!.Value.StartLine);
        Assert.Equal("One | 1", firstRow.Literal);

        var secondRow = table.Children[1];
        Assert.Equal(MarkdownSyntaxKind.TableRow, secondRow.Kind);
        Assert.NotNull(secondRow.SourceSpan);
        Assert.Equal(2, secondRow.SourceSpan!.Value.StartLine);
        Assert.Equal("Two | 2", secondRow.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Fenced_Code_Block_Structure() {
        var markdown = """
```csharp
Console.WriteLine("hi");
```
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var code = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.CodeBlock, code.Kind);
        Assert.NotNull(code.SourceSpan);
        Assert.Equal(1, code.SourceSpan!.Value.StartLine);
        Assert.Equal(3, code.SourceSpan!.Value.EndLine);
        Assert.Equal(2, code.Children.Count);

        var info = code.Children[0];
        Assert.Equal(MarkdownSyntaxKind.CodeFenceInfo, info.Kind);
        Assert.NotNull(info.SourceSpan);
        Assert.Equal(1, info.SourceSpan!.Value.StartLine);
        Assert.Equal("csharp", info.Literal);

        var content = code.Children[1];
        Assert.Equal(MarkdownSyntaxKind.CodeContent, content.Kind);
        Assert.NotNull(content.SourceSpan);
        Assert.Equal(2, content.SourceSpan!.Value.StartLine);
        Assert.Equal(2, content.SourceSpan!.Value.EndLine);
        Assert.Equal("Console.WriteLine(\"hi\");", content.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Preserves_Raw_Fence_InfoString_Literal() {
        var markdown = """
```json title="chart"
{"value":1}
```
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var code = Assert.Single(result.SyntaxTree.Children);
        var info = code.Children[0];

        Assert.Equal(MarkdownSyntaxKind.CodeFenceInfo, info.Kind);
        Assert.Equal("json title=\"chart\"", info.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Indented_Code_Block_Structure() {
        var markdown = """
    line 1
    line 2
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var code = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.CodeBlock, code.Kind);
        Assert.Single(code.Children);

        var content = code.Children[0];
        Assert.Equal(MarkdownSyntaxKind.CodeContent, content.Kind);
        Assert.NotNull(content.SourceSpan);
        Assert.Equal(1, content.SourceSpan!.Value.StartLine);
        Assert.Equal(2, content.SourceSpan!.Value.EndLine);
        Assert.Equal("line 1\nline 2", content.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Image_Structure() {
        var markdown = """
![Alt text](https://example.com/image.png "Image title")
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var image = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Image, image.Kind);
        Assert.NotNull(image.SourceSpan);
        Assert.Equal(1, image.SourceSpan!.Value.StartLine);
        Assert.Equal(1, image.SourceSpan!.Value.EndLine);
        Assert.Equal(3, image.Children.Count);

        var alt = image.Children[0];
        Assert.Equal(MarkdownSyntaxKind.ImageAlt, alt.Kind);
        Assert.Equal("Alt text", alt.Literal);

        var source = image.Children[1];
        Assert.Equal(MarkdownSyntaxKind.ImageSource, source.Kind);
        Assert.Equal("https://example.com/image.png", source.Literal);

        var title = image.Children[2];
        Assert.Equal(MarkdownSyntaxKind.ImageTitle, title.Kind);
        Assert.Equal("Image title", title.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Linked_Image_Block_Metadata() {
        var markdown = """
[![Alt text](https://example.com/image.png "Image title")](https://example.com/docs "Link title")
_Caption_
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var image = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Image, image.Kind);

        Assert.Collection(image.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageAlt, node.Kind);
                Assert.Equal("Alt text", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageSource, node.Kind);
                Assert.Equal("https://example.com/image.png", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTarget, node.Kind);
                Assert.Equal("https://example.com/docs", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTitle, node.Kind);
                Assert.Equal("Link title", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageTitle, node.Kind);
                Assert.Equal("Image title", node.Literal);
            });
    }

    [Fact]
    public void HtmlImported_Image_SyntaxNode_Captures_Linked_Html_Metadata() {
        const string html = """
<figure>
  <a href="/docs/hero" title="Hero page" target="_blank" rel="nofollow sponsored">
    <img src="/img/hero.png" alt="Hero" title="View hero" />
  </a>
  <figcaption>Hero image</figcaption>
</figure>
""";

        var document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        var syntax = ((ISyntaxMarkdownBlock)image).BuildSyntaxNode(null);

        Assert.Collection(syntax.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageAlt, node.Kind);
                Assert.Equal("Hero", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageSource, node.Kind);
                Assert.Equal("https://example.com/img/hero.png", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTarget, node.Kind);
                Assert.Equal("https://example.com/docs/hero", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTitle, node.Kind);
                Assert.Equal("Hero page", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkHtmlTarget, node.Kind);
                Assert.Equal("_blank", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkHtmlRel, node.Kind);
                Assert.Equal("nofollow sponsored", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageTitle, node.Kind);
                Assert.Equal("View hero", node.Literal);
            });
    }

    [Fact]
    public void HtmlImported_Wrapped_Picture_SyntaxNode_Captures_Linked_Html_Metadata() {
        const string html = """
<figure>
  <a href="/docs/hero" title="Hero page" target="_blank" rel="nofollow sponsored">
    <div class="media-wrap">
      <picture>
        <source srcset="/img/hero.webp" type="image/webp" />
        <img src="/img/hero.png" alt="Hero" title="View hero" />
      </picture>
    </div>
  </a>
  <figcaption>Hero image</figcaption>
</figure>
""";

        var document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        var syntax = ((ISyntaxMarkdownBlock)image).BuildSyntaxNode(null);

        Assert.Collection(syntax.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageAlt, node.Kind);
                Assert.Equal("Hero", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageSource, node.Kind);
                Assert.Equal("https://example.com/img/hero.webp", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTarget, node.Kind);
                Assert.Equal("https://example.com/docs/hero", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTitle, node.Kind);
                Assert.Equal("Hero page", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkHtmlTarget, node.Kind);
                Assert.Equal("_blank", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkHtmlRel, node.Kind);
                Assert.Equal("nofollow sponsored", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageTitle, node.Kind);
                Assert.Equal("View hero", node.Literal);
            });
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Front_Matter_Block() {
        var markdown = """
--- 
title: Sample
---
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var frontMatter = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.FrontMatter, frontMatter.Kind);
        Assert.NotNull(frontMatter.SourceSpan);
        Assert.Equal(1, frontMatter.SourceSpan!.Value.StartLine);
        Assert.Equal(3, frontMatter.SourceSpan!.Value.EndLine);
        Assert.Equal("---\ntitle: Sample\n---", frontMatter.Literal!.Replace("\r\n", "\n"));
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Html_Comment_Block() {
        const string markdown = "<!-- keep me -->";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var comment = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.HtmlComment, comment.Kind);
        Assert.NotNull(comment.SourceSpan);
        Assert.Equal(1, comment.SourceSpan!.Value.StartLine);
        Assert.Equal(1, comment.SourceSpan!.Value.EndLine);
        Assert.Equal(markdown, comment.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Html_Raw_Block() {
        const string markdown = "<div class=\"note\">Hello</div>";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var rawHtml = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.HtmlRaw, rawHtml.Kind);
        Assert.NotNull(rawHtml.SourceSpan);
        Assert.Equal(1, rawHtml.SourceSpan!.Value.StartLine);
        Assert.Equal(1, rawHtml.SourceSpan!.Value.EndLine);
        Assert.Equal(markdown, rawHtml.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Toc_Placeholder_Block() {
        const string markdown = "[TOC]";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var tocPlaceholder = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.TocPlaceholder, tocPlaceholder.Kind);
        Assert.NotNull(tocPlaceholder.SourceSpan);
        Assert.Equal(1, tocPlaceholder.SourceSpan!.Value.StartLine);
        Assert.Equal(1, tocPlaceholder.SourceSpan!.Value.EndLine);
        Assert.Null(tocPlaceholder.Literal);
        Assert.Empty(tocPlaceholder.Children);
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
        Assert.Equal(MarkdownSyntaxKind.InlineText, titleNode!.Kind);
        Assert.Equal("Title", titleNode.Literal);

        var leadNode = result.SyntaxTree.FindDeepestNodeAtLine(3);
        Assert.NotNull(leadNode);
        Assert.Equal(MarkdownSyntaxKind.InlineText, leadNode!.Kind);
        Assert.Equal("lead continued", leadNode.Literal);

        var quoteNode = result.SyntaxTree.FindDeepestNodeAtLine(6);
        Assert.NotNull(quoteNode);
        Assert.Equal(MarkdownSyntaxKind.InlineText, quoteNode!.Kind);
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

        Assert.Equal(new[] { MarkdownSyntaxKind.Document, MarkdownSyntaxKind.Paragraph, MarkdownSyntaxKind.InlineText }, kinds);
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
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineText
        }, path);

        Assert.Empty(result.SyntaxTree.FindNodePathAtLine(99));
    }

    [Fact]
    public void ParseWithSyntaxTree_Finds_Nearest_Block_By_Line() {
        var markdown = """
```csharp
Console.WriteLine();
```

![Alt](image.png "Image title")
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var codeDeepest = result.FindDeepestNodeAtLine(1);
        Assert.NotNull(codeDeepest);
        Assert.Equal(MarkdownSyntaxKind.CodeFenceInfo, codeDeepest!.Kind);

        var codeBlock = result.FindNearestBlockAtLine(1);
        Assert.NotNull(codeBlock);
        Assert.Equal(MarkdownSyntaxKind.CodeBlock, codeBlock!.Kind);

        var imageDeepest = result.FindDeepestNodeAtLine(5);
        Assert.NotNull(imageDeepest);
        Assert.Equal(MarkdownSyntaxKind.ImageAlt, imageDeepest!.Kind);

        var imageBlock = result.FindNearestBlockAtLine(5);
        Assert.NotNull(imageBlock);
        Assert.Equal(MarkdownSyntaxKind.Image, imageBlock!.Kind);

        Assert.Null(result.FindNearestBlockAtLine(99));
    }

    [Fact]
    public void ParseWithSyntaxTree_Result_Provides_Line_Lookup_Helpers() {
        var markdown = """
# Title

Paragraph
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var deepest = result.FindDeepestNodeAtLine(3);
        Assert.NotNull(deepest);
        Assert.Equal(MarkdownSyntaxKind.InlineText, deepest!.Kind);
        Assert.Equal("Paragraph", deepest.Literal);

        var path = result.FindNodePathAtLine(1).Select(node => node.Kind).ToArray();
        Assert.Equal(new[] { MarkdownSyntaxKind.Document, MarkdownSyntaxKind.Heading, MarkdownSyntaxKind.HeadingText, MarkdownSyntaxKind.InlineText }, path);

        var nearest = result.FindNearestBlockAtLine(1);
        Assert.NotNull(nearest);
        Assert.Equal(MarkdownSyntaxKind.Heading, nearest!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTree_Finds_Deepest_Node_By_Span() {
        var markdown = """
> [!TIP] Title
> - item
>   continued
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var deepest = result.FindDeepestNodeContainingSpan(new MarkdownSourceSpan(2, 3));
        Assert.NotNull(deepest);
        Assert.Equal(MarkdownSyntaxKind.InlineText, deepest!.Kind);
        Assert.Equal("item continued", deepest.Literal);

        var path = result.FindNodePathContainingSpan(new MarkdownSourceSpan(2, 3)).Select(node => node.Kind).ToArray();
        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.Callout,
            MarkdownSyntaxKind.UnorderedList,
            MarkdownSyntaxKind.ListItem,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineText
        }, path);

        Assert.Null(result.FindDeepestNodeContainingSpan(new MarkdownSourceSpan(50, 51)));
        Assert.Empty(result.FindNodePathContainingSpan(new MarkdownSourceSpan(50, 51)));
    }

    [Fact]
    public void ParseWithSyntaxTree_Finds_Deepest_Node_By_Overlapping_Span() {
        var markdown = """
# Title

Paragraph text
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var deepest = result.FindDeepestNodeOverlappingSpan(new MarkdownSourceSpan(1, 2));
        Assert.NotNull(deepest);
        Assert.Equal(MarkdownSyntaxKind.InlineText, deepest!.Kind);
        Assert.Equal("Title", deepest.Literal);

        var path = result.FindNodePathOverlappingSpan(new MarkdownSourceSpan(2, 3)).Select(node => node.Kind).ToArray();
        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineText
        }, path);

        Assert.Null(result.FindDeepestNodeOverlappingSpan(new MarkdownSourceSpan(50, 51)));
        Assert.Empty(result.FindNodePathOverlappingSpan(new MarkdownSourceSpan(50, 51)));
    }

    [Fact]
    public void ParseWithSyntaxTree_Finds_Nearest_Block_By_Span() {
        var markdown = """
```csharp
Console.WriteLine();
```

![Alt](image.png "Image title")
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var codeBlock = result.FindNearestBlockContainingSpan(new MarkdownSourceSpan(1, 1));
        Assert.NotNull(codeBlock);
        Assert.Equal(MarkdownSyntaxKind.CodeBlock, codeBlock!.Kind);

        var imageBlock = result.FindNearestBlockOverlappingSpan(new MarkdownSourceSpan(5, 5));
        Assert.NotNull(imageBlock);
        Assert.Equal(MarkdownSyntaxKind.Image, imageBlock!.Kind);

        Assert.Null(result.FindNearestBlockContainingSpan(new MarkdownSourceSpan(50, 51)));
        Assert.Null(result.FindNearestBlockOverlappingSpan(new MarkdownSourceSpan(50, 51)));
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

    private sealed class RewriteFirstParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var rewritten = MarkdownDoc.Create();
            if (document.DocumentHeader != null) {
                rewritten.Add(document.DocumentHeader);
            }

            for (var i = 0; i < document.Blocks.Count; i++) {
                if (i == 0) {
                    rewritten.Add(new ParagraphBlock(new InlineSequence().Text(text)));
                } else {
                    rewritten.Add(document.Blocks[i]);
                }
            }

            return rewritten;
        }
    }

    private sealed class RewriteDefinitionListDefinitionsTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var rewritten = MarkdownDoc.Create();
            if (document.DocumentHeader != null) {
                rewritten.Add(document.DocumentHeader);
            }

            foreach (var block in document.Blocks) {
                if (block is not DefinitionListBlock definitionList) {
                    rewritten.Add(block);
                    continue;
                }

                var rebuilt = new DefinitionListBlock();
                foreach (var entry in definitionList.Entries) {
                    rebuilt.AddEntry(new DefinitionListEntry(
                        entry.Term,
                        new[] { new ParagraphBlock(new InlineSequence().Text(text)) }));
                }

                rewritten.Add(rebuilt);
            }

            return rewritten;
        }
    }

    private sealed class RewriteNestedParagraphsTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            MarkdownDocumentBlockRewriter.RewriteDocument(document, block =>
                block is ParagraphBlock
                    ? new ParagraphBlock(new InlineSequence().Text(text))
                    : block);
            return document;
        }
    }
}
