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
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 42), diagnostic.AffectedSourceSpan);
        Assert.Equal("Document > Paragraph", diagnostic.AffectedOriginalBlockPath);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 42), diagnostic.AffectedOriginalBlockSpan);
        Assert.Equal("Document > Paragraph", diagnostic.AffectedFinalBlockPath);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 42), diagnostic.AffectedFinalBlockSpan);
        Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(2, result.FinalSyntaxTree.Children.Count);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 42), result.FinalSyntaxTree.Children[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 42), result.FinalSyntaxTree.Children[1].SourceSpan);
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
    public void ParseWithSyntaxTreeAndDiagnostics_Detaches_Original_Syntax_Associations_When_Transform_Replaces_Document() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteFirstParagraphTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("hello", options);

        Assert.Null(result.SyntaxTree.AssociatedObject);
        Assert.Null(Assert.Single(result.SyntaxTree.Children).AssociatedObject);
        Assert.Same(result.Document, result.FinalSyntaxTree.AssociatedObject);
        Assert.IsType<ParagraphBlock>(Assert.Single(result.FinalSyntaxTree.Children).AssociatedObject);
    }

    [Fact]
    public void ParseWithSyntaxTree_Associates_Final_Nested_List_Syntax_To_Live_Semantic_Objects() {
        var markdown = """
- outer
  - inner

    tail
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var semanticOuterList = Assert.IsType<UnorderedListBlock>(Assert.Single(result.Document.Blocks));
        var semanticOuterItem = Assert.IsType<ListItem>(Assert.Single(semanticOuterList.Items));
        var semanticNestedList = Assert.IsType<UnorderedListBlock>(Assert.Single(semanticOuterItem.ChildBlocks));
        var semanticNestedItem = Assert.IsType<ListItem>(Assert.Single(semanticNestedList.Items));
        var semanticNestedParagraphs = semanticNestedItem.ParagraphBlocks;

        var finalOuterList = Assert.Single(result.FinalSyntaxTree.Children);
        var finalOuterItem = Assert.Single(finalOuterList.Children);
        var finalNestedList = Assert.Single(finalOuterItem.Children, node => node.Kind == MarkdownSyntaxKind.UnorderedList);
        var finalNestedItem = Assert.Single(finalNestedList.Children);
        var finalNestedParagraphs = finalNestedItem.Children.Where(node => node.Kind == MarkdownSyntaxKind.Paragraph).ToArray();

        Assert.Same(semanticOuterList, finalOuterList.AssociatedObject);
        Assert.Same(semanticOuterItem, finalOuterItem.AssociatedObject);
        Assert.Same(semanticNestedList, finalNestedList.AssociatedObject);
        Assert.Same(semanticNestedItem, finalNestedItem.AssociatedObject);
        Assert.Equal(2, finalNestedParagraphs.Length);
        Assert.Same(semanticNestedParagraphs[0], finalNestedParagraphs[0].AssociatedObject);
        Assert.Same(semanticNestedParagraphs[1], finalNestedParagraphs[1].AssociatedObject);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Transform_SourceSpans_When_Original_Syntax_Has_Reference_Definitions() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteFirstParagraphTransform("rewritten"));
        var markdown = """
[hero]: https://example.com/docs "Docs title"

[hero]
""";

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

        var diagnostic = Assert.Single(result.TransformDiagnostics);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 6), diagnostic.AffectedSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 6), Assert.Single(result.FinalSyntaxTree.Children).SourceSpan);
        Assert.Equal("rewritten", Assert.Single(result.FinalSyntaxTree.Children).Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Aggregate_SourceSpans_For_Adjacent_Block_Rewrites() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteFirstTwoParagraphsTransform("first rewritten", "second rewritten"));
        var markdown = """
alpha

beta

gamma
""";

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

        var diagnostic = Assert.Single(result.TransformDiagnostics);
        Assert.Equal(0, diagnostic.ChangedBlockStartBefore);
        Assert.Equal(2, diagnostic.ChangedBlockCountBefore);
        Assert.Equal(0, diagnostic.ChangedBlockStartAfter);
        Assert.Equal(2, diagnostic.ChangedBlockCountAfter);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), diagnostic.AffectedSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), result.FinalSyntaxTree.Children[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), result.FinalSyntaxTree.Children[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 5), result.FinalSyntaxTree.Children[2].SourceSpan);
        Assert.Equal("gamma", result.FinalSyntaxTree.Children[2].Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Aggregate_SourceSpans_For_Adjacent_Block_Merges() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsTransform("merged"));
        var markdown = """
alpha

beta

gamma
""";

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

        var diagnostic = Assert.Single(result.TransformDiagnostics);
        Assert.Equal(0, diagnostic.ChangedBlockStartBefore);
        Assert.Equal(2, diagnostic.ChangedBlockCountBefore);
        Assert.Equal(0, diagnostic.ChangedBlockStartAfter);
        Assert.Equal(1, diagnostic.ChangedBlockCountAfter);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), diagnostic.AffectedSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), result.FinalSyntaxTree.Children[0].SourceSpan);
        Assert.Equal("merged", result.FinalSyntaxTree.Children[0].Literal);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 5), result.FinalSyntaxTree.Children[1].SourceSpan);
        Assert.Equal("gamma", result.FinalSyntaxTree.Children[1].Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_SourceProvenance_Across_Chained_Block_Transforms() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteFirstTwoParagraphsTransform("first rewritten", "second rewritten"));
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsTransform("merged"));
        var markdown = """
alpha

beta

gamma
""";

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

        Assert.Equal(2, result.TransformDiagnostics.Count);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), result.TransformDiagnostics[0].AffectedSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), result.TransformDiagnostics[1].AffectedSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), result.FinalSyntaxTree.Children[0].SourceSpan);
        Assert.Equal("merged", result.FinalSyntaxTree.Children[0].Literal);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 5), result.FinalSyntaxTree.Children[1].SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Keeps_InsertOnly_Blocks_Without_SourceProvenance_Across_Later_Rewrites() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new AppendParagraphTransform("tail"));
        options.DocumentTransforms.Add(new RewriteSecondParagraphTransform("tail rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("hello", options);

        Assert.Equal(2, result.TransformDiagnostics.Count);
        Assert.Null(result.TransformDiagnostics[0].AffectedSourceSpan);
        Assert.Null(result.TransformDiagnostics[1].AffectedSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 5), result.FinalSyntaxTree.Children[0].SourceSpan);
        Assert.Null(result.FinalSyntaxTree.Children[1].SourceSpan);
        Assert.Equal("tail rewritten", result.FinalSyntaxTree.Children[1].Literal);
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
        Assert.Equal(2, link.Children.Count);
        Assert.Equal(MarkdownSyntaxKind.InlineText, link.Children[0].Kind);
        Assert.Equal("docs", link.Children[0].Literal);
        Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, link.Children[1].Kind);
        Assert.Equal("https://example.com", link.Children[1].Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 21, 1, 39), link.Children[1].SourceSpan);

        var code = paragraph.Children[5];
        Assert.Equal(MarkdownSyntaxKind.InlineCodeSpan, code.Kind);
        Assert.Equal("code", code.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Inline_Link_Metadata_Nodes() {
        const string markdown = "[docs](https://example.com \"Example title\")";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var link = Assert.Single(paragraph.Children);

        Assert.Equal(MarkdownSyntaxKind.InlineLink, link.Kind);
        Assert.Equal("https://example.com", link.Literal);
        Assert.Collection(link.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineText, node.Kind);
                Assert.Equal("docs", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, node.Kind);
                Assert.Equal("https://example.com", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 26), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTitle, node.Kind);
                Assert.Equal("Example title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 29, 1, 41), node.SourceSpan);
            });
    }

    [Fact]
    public void ParseWithSyntaxTree_Uses_Custom_Inline_Parser_Extensions_With_Nested_Ast_And_SourceSpans() {
        const string markdown = "Lead {{**Bold** core}} tail";
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension("double-brace", TryParseDoubleBraceInline));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        var paragraph = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.Unknown,
            MarkdownSyntaxKind.InlineText
        }, paragraph.Children.Select(node => node.Kind).ToArray());

        var custom = paragraph.Children[1];
        Assert.Equal("double-brace", custom.CustomKind);
        Assert.Equal("{{**Bold** core}}", custom.Literal);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineStrong,
            MarkdownSyntaxKind.InlineText
        }, custom.Children.Select(node => node.Kind).ToArray());

        var customStart = markdown.IndexOf("{{", StringComparison.Ordinal) + 1;
        var customEnd = markdown.IndexOf("}}", StringComparison.Ordinal) + 2;
        Assert.Equal(customStart, custom.SourceSpan!.Value.StartColumn);
        Assert.Equal(customEnd, custom.SourceSpan!.Value.EndColumn);
        Assert.Equal(custom.SourceSpan, ((DoubleBraceInline)custom.AssociatedObject!).SourceSpan);

        var nestedStrong = custom.Children[0];
        Assert.Equal(10, nestedStrong.SourceSpan!.Value.StartColumn);
        Assert.Equal(13, nestedStrong.SourceSpan!.Value.EndColumn);
        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(1, 18)!.Kind);
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
        Assert.Equal(new MarkdownSourceSpan(1, 21, 1, 39), link.Children[1].SourceSpan);

        var code = paragraph.Children[5];
        Assert.Equal(46, code.SourceSpan!.Value.StartColumn);
        Assert.Equal(51, code.SourceSpan!.Value.EndColumn);

        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(1, 8)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, result.FindDeepestNodeAtPosition(1, 30)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.InlineCodeSpan, result.FindDeepestNodeAtPosition(1, 48)!.Kind);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineLink,
            MarkdownSyntaxKind.InlineLinkTarget
        }, result.FindNodePathAtPosition(1, 30).Select(node => node.Kind).ToArray());
        Assert.Equal(MarkdownSyntaxKind.Paragraph, result.FindNearestBlockAtPosition(1, 48)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTree_Associates_SequenceInline_Syntax_To_Wrapper_Objects() {
        const string markdown = "Use **bold** and *emphasis*.";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var paragraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        var paragraphSyntax = Assert.Single(result.FinalSyntaxTree.Children);
        var boldInline = Assert.IsType<BoldSequenceInline>(paragraphBlock.Inlines.Nodes[1]);
        var italicInline = Assert.IsType<ItalicSequenceInline>(paragraphBlock.Inlines.Nodes[3]);

        var strongSyntax = paragraphSyntax.Children[1];
        var emphasisSyntax = paragraphSyntax.Children[3];

        Assert.Same(boldInline, strongSyntax.AssociatedObject);
        Assert.Equal(boldInline.SourceSpan, strongSyntax.SourceSpan);
        Assert.Same(italicInline, emphasisSyntax.AssociatedObject);
        Assert.Equal(italicInline.SourceSpan, emphasisSyntax.SourceSpan);
        Assert.IsNotType<InlineSequence>(strongSyntax.AssociatedObject);
        Assert.IsNotType<InlineSequence>(emphasisSyntax.AssociatedObject);
    }

    [Fact]
    public void Table_RowInlines_Reuse_Custom_Inline_Parser_Extensions_When_Table_Clones_Reader_Options() {
        const string markdown = """
| Name |
| --- |
| {{Cell}} |
""";
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension("double-brace", TryParseDoubleBraceInline));

        var document = MarkdownReader.Parse(markdown, options);

        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));
        var cellInlines = Assert.Single(Assert.Single(table.RowInlines));
        var inline = Assert.IsType<DoubleBraceInline>(Assert.Single(cellInlines.Nodes));
        Assert.Equal("Cell", InlinePlainText.Extract(inline.Inlines));
    }

    [Fact]
    public void Custom_Inline_Can_Render_Html_With_Public_Html_Options_Context() {
        const string markdown = "Lead {{core}} tail";
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension("double-brace", TryParseDoubleBraceInline));

        var document = MarkdownReader.Parse(markdown, options);
        var html = document.ToHtmlFragment(new HtmlOptions {
            Kind = HtmlKind.Fragment,
            Title = "inline-title"
        });

        Assert.Contains("data-inline=\"double-brace\"", html, StringComparison.Ordinal);
        Assert.Contains("data-title=\"inline-title\"", html, StringComparison.Ordinal);
        Assert.Contains(">core<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Assigns_Parent_Sibling_And_AssociatedObject_Metadata() {
        const string markdown = "Use **bold** [docs](https://example.com) and `code`.";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Same(result.Document, result.SyntaxTree.AssociatedObject);
        Assert.Null(result.SyntaxTree.Parent);
        Assert.Same(result.SyntaxTree, result.SyntaxTree.Root);

        var paragraph = Assert.Single(result.SyntaxTree.Children);
        Assert.Same(result.SyntaxTree, paragraph.Parent);
        Assert.Equal(0, paragraph.IndexInParent);
        Assert.Null(paragraph.PreviousSibling);
        Assert.Null(paragraph.NextSibling);
        Assert.IsType<ParagraphBlock>(paragraph.AssociatedObject);

        var link = paragraph.Children[3];
        Assert.Same(paragraph, link.Parent);
        Assert.Equal(3, link.IndexInParent);
        Assert.Equal(MarkdownSyntaxKind.InlineText, link.PreviousSibling!.Kind);
        Assert.Equal(MarkdownSyntaxKind.InlineText, link.NextSibling!.Kind);
        Assert.IsType<LinkInline>(link.AssociatedObject);
        Assert.Equal(new[] { MarkdownSyntaxKind.Paragraph, MarkdownSyntaxKind.Document }, link.Ancestors().Select(node => node.Kind).ToArray());
        Assert.Equal(new[] { MarkdownSyntaxKind.InlineLink, MarkdownSyntaxKind.Paragraph, MarkdownSyntaxKind.Document }, link.AncestorsAndSelf().Select(node => node.Kind).ToArray());
        Assert.Same(result.SyntaxTree, link.Root);
    }

    [Fact]
    public void ParseWithSyntaxTree_Assigns_ObjectModel_Parents_Siblings_And_SourceSpans() {
        const string markdown = """
# Title

- first [link](https://example.com)
- second
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var document = result.Document;
        var heading = Assert.IsType<HeadingBlock>(document.Blocks[0]);
        var list = Assert.IsType<UnorderedListBlock>(document.Blocks[1]);
        var firstItem = Assert.IsType<ListItem>(list.Items[0]);
        var secondItem = Assert.IsType<ListItem>(list.Items[1]);
        var headingText = Assert.IsType<TextRun>(Assert.Single(heading.Inlines.Nodes));
        var link = Assert.Single(firstItem.Content.Nodes.OfType<LinkInline>());

        Assert.Null(document.Parent);
        Assert.Same(document, heading.Parent);
        Assert.Same(document, list.Parent);
        Assert.Equal(0, heading.IndexInParent);
        Assert.Equal(1, list.IndexInParent);
        Assert.Null(heading.PreviousSibling);
        Assert.Same(list, heading.NextSibling);
        Assert.Same(heading, list.PreviousSibling);

        Assert.Same(heading, heading.Inlines.Parent);
        Assert.Same(heading.Inlines, headingText.Parent);
        Assert.Same(list, firstItem.Parent);
        Assert.Same(list, secondItem.Parent);
        Assert.Equal(0, firstItem.IndexInParent);
        Assert.Equal(1, secondItem.IndexInParent);
        Assert.Same(secondItem, firstItem.NextSibling);
        Assert.Same(firstItem, secondItem.PreviousSibling);
        var firstParagraph = Assert.IsType<ParagraphBlock>(firstItem.BlockChildren[0]);
        Assert.Same(firstItem, firstParagraph.Parent);
        Assert.Same(firstParagraph, firstItem.Content.Parent);
        Assert.Same(firstItem.Content, link.Parent);

        Assert.Same(document, link.Document);
        Assert.Same(document, link.Root);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 4, 8), document.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 7), heading.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 7), heading.Inlines.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 7), headingText.SourceSpan);
    }

    [Fact]
    public void FluentDocument_Assigns_ObjectModel_Parents_Without_SyntaxTree() {
        var document = MarkdownDoc.Create()
            .H1("Title")
            .Ul(list => {
                list.Item("first");
                list.Item("second");
            });

        var heading = Assert.IsType<HeadingBlock>(document.Blocks[0]);
        var list = Assert.IsType<UnorderedListBlock>(document.Blocks[1]);
        var firstItem = Assert.IsType<ListItem>(list.Items[0]);
        var secondItem = Assert.IsType<ListItem>(list.Items[1]);
        var firstText = Assert.IsType<TextRun>(Assert.Single(firstItem.Content.Nodes));

        Assert.Same(document, heading.Parent);
        Assert.Same(document, list.Parent);
        Assert.Same(heading, heading.Inlines.Parent);
        Assert.Same(list, firstItem.Parent);
        Assert.Same(list, secondItem.Parent);
        var firstParagraph = Assert.IsType<ParagraphBlock>(firstItem.BlockChildren[0]);
        Assert.Same(firstItem, firstParagraph.Parent);
        Assert.Same(firstParagraph, firstItem.Content.Parent);
        Assert.Same(firstItem.Content, firstText.Parent);
        Assert.Null(heading.SourceSpan);
        Assert.Equal(new MarkdownObject[] { heading, list }, document.ChildObjects.ToArray());
        Assert.Equal(new MarkdownObject[] { firstItem, secondItem }, list.ChildObjects.ToArray());
        Assert.Equal(new MarkdownObject[] { firstItem, list, document }, firstItem.AncestorsAndSelf().ToArray());
    }

    [Fact]
    public void ListItem_ParagraphBlocks_And_TableCells_Are_Stable_Owned_Nodes() {
        const string markdown = """
- first paragraph

  second paragraph

| Name | Value |
| --- | --- |
| One | 1 |
""";

        var document = MarkdownReader.Parse(markdown);
        var list = Assert.IsType<UnorderedListBlock>(document.Blocks[0]);
        var item = Assert.Single(list.Items);
        var table = Assert.IsType<TableBlock>(document.Blocks[1]);

        var firstParagraphRead1 = item.ParagraphBlocks[0];
        var firstParagraphRead2 = item.ParagraphBlocks[0];
        var secondParagraph = item.ParagraphBlocks[1];
        var blockChildrenRead1 = item.BlockChildren;
        var blockChildrenRead2 = item.BlockChildren;

        Assert.Same(firstParagraphRead1, firstParagraphRead2);
        Assert.Same(blockChildrenRead1[0], blockChildrenRead2[0]);
        Assert.Same(item, firstParagraphRead1.Parent);
        Assert.Same(item, secondParagraph.Parent);
        Assert.Same(firstParagraphRead1, item.Content.Parent);
        Assert.Same(secondParagraph, item.AdditionalParagraphs[0].Parent);

        var headerRead1 = table.HeaderCells[0];
        var headerRead2 = table.HeaderCells[0];
        var bodyRead1 = table.RowCells[0][1];
        var bodyRead2 = table.GetCell(0, 1);

        Assert.Same(headerRead1, headerRead2);
        Assert.Same(bodyRead1, bodyRead2);
        Assert.Same(table, headerRead1.Parent);
        Assert.Same(table, bodyRead1.Parent);
        Assert.All(table.EnumerateCells(), cell => Assert.Same(table, cell.Parent));
    }

    [Fact]
    public void ParseWithSyntaxTree_Associates_ListItem_Paragraph_Syntax_To_ParagraphBlocks() {
        const string markdown = """
- first paragraph

  second paragraph
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(result.Document.Blocks));
        var item = Assert.Single(list.Items);
        var listSyntax = Assert.Single(result.FinalSyntaxTree.Children);
        var itemSyntax = Assert.Single(listSyntax.Children);

        Assert.Equal(2, item.ParagraphBlocks.Count);
        Assert.Equal(2, itemSyntax.Children.Count);
        Assert.All(itemSyntax.Children, child => Assert.Equal(MarkdownSyntaxKind.Paragraph, child.Kind));
        Assert.Same(item.ParagraphBlocks[0], itemSyntax.Children[0].AssociatedObject);
        Assert.Same(item.ParagraphBlocks[1], itemSyntax.Children[1].AssociatedObject);
        Assert.Equal(item.ParagraphBlocks[0].SourceSpan, itemSyntax.Children[0].SourceSpan);
        Assert.Equal(item.ParagraphBlocks[1].SourceSpan, itemSyntax.Children[1].SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTree_Assigns_SourceSpans_To_TableCell_Ast_Objects() {
        const string markdown = """
| Name | Value |
| --- | --- |
| One | 1 |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var table = Assert.IsType<TableBlock>(Assert.Single(result.Document.Blocks));

        var header = table.GetHeaderCell(0);
        var body = table.GetCell(0, 1);

        Assert.NotNull(header);
        Assert.NotNull(body);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 11), table.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 6), header!.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 9, 3, 9), body!.SourceSpan);
    }

    [Fact]
    public void MarkdownVisitor_Walks_Public_ObjectTree_In_DepthFirst_Order() {
        var document = MarkdownReader.Parse("""
# Title

- first

| Name | Value |
| --- | --- |
| One | 1 |
""");

        var visitor = new CollectingMarkdownVisitor();
        document.Accept(visitor);

        Assert.Equal(new[] {
            "MarkdownDoc",
            "HeadingBlock",
            "InlineSequence",
            "TextRun",
            "UnorderedListBlock",
            "ListItem",
            "ParagraphBlock",
            "InlineSequence",
            "TextRun",
            "TableBlock",
            "TableCell",
            "ParagraphBlock",
            "InlineSequence",
            "TextRun",
            "TableCell",
            "ParagraphBlock",
            "InlineSequence",
            "TextRun",
            "TableCell",
            "ParagraphBlock",
            "InlineSequence",
            "TextRun",
            "TableCell",
            "ParagraphBlock",
            "InlineSequence",
            "TextRun"
        }, visitor.NodeKinds);
    }

    [Fact]
    public void MarkdownRewriter_Rewrites_Nested_Block_Content_And_Rebinds_Parents() {
        var document = MarkdownReader.Parse("""
> before

- item
""");

        document.Rewrite(new ReplaceParagraphRewriter("after"));

        var quote = Assert.IsType<QuoteBlock>(document.Blocks[0]);
        var quoteParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        Assert.Equal("after", quoteParagraph.Inlines.RenderMarkdown());
        Assert.Same(quote, quoteParagraph.Parent);
        Assert.Same(quoteParagraph, quoteParagraph.Inlines.Parent);

        var list = Assert.IsType<UnorderedListBlock>(document.Blocks[1]);
        var itemParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(list.Items[0].BlockChildren));
        Assert.Equal("after", itemParagraph.Inlines.RenderMarkdown());
        Assert.Same(list.Items[0], itemParagraph.Parent);
        Assert.Same(itemParagraph, itemParagraph.Inlines.Parent);
    }

    [Fact]
    public void MarkdownSourceSpan_Uses_ColumnAware_Equality_Containment_And_Overlap() {
        var outer = new MarkdownSourceSpan(3, 5, 3, 20);
        var inner = new MarkdownSourceSpan(3, 8, 3, 12);
        var disjointSameLine = new MarkdownSourceSpan(3, 21, 3, 24);
        var sameLinesDifferentColumns = new MarkdownSourceSpan(3, 1, 3, 4);

        Assert.NotEqual(outer, sameLinesDifferentColumns);
        Assert.True(outer.Contains(inner));
        Assert.False(outer.Contains(disjointSameLine));
        Assert.False(outer.Overlaps(disjointSameLine));
        Assert.True(outer.Overlaps(new MarkdownSourceSpan(3, 20, 3, 24)));
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
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 9), image.Children[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 12, 1, 20), image.Children[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 23, 1, 27), image.Children[2].SourceSpan);
        Assert.Equal(MarkdownSyntaxKind.ImageSource, result.FindDeepestNodeAtPosition(1, 15)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ImageTitle, result.FindDeepestNodeAtPosition(1, 24)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Inline_Image_Link_Metadata_Nodes_And_Spans() {
        const string markdown = "See [![Alt text](https://example.com/image.png \"Image title\")](https://example.com/docs \"Link title\")";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var paragraph = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineImageLink
        }, paragraph.Children.Select(node => node.Kind).ToArray());

        var imageLink = paragraph.Children[1];
        Assert.Equal("https://example.com/docs", imageLink.Literal);
        Assert.Collection(imageLink.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageAlt, node.Kind);
                Assert.Equal("Alt text", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 15), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageSource, node.Kind);
                Assert.Equal("https://example.com/image.png", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 18, 1, 46), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTarget, node.Kind);
                Assert.Equal("https://example.com/docs", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 64, 1, 87), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTitle, node.Kind);
                Assert.Equal("Link title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 90, 1, 99), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageTitle, node.Kind);
                Assert.Equal("Image title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 49, 1, 59), node.SourceSpan);
            });

        Assert.Equal(MarkdownSyntaxKind.ImageSource, result.FindDeepestNodeAtPosition(1, 25)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ImageTitle, result.FindDeepestNodeAtPosition(1, 52)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ImageLinkTitle, result.FindDeepestNodeAtPosition(1, 92)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTree_Preserves_Raw_Image_Alt_Syntax_While_Rendering_Plain_Alt_Text() {
        const string markdown = "Lead ![foo *bar*](train.jpg \"train & tracks\")";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var paragraph = Assert.Single(result.SyntaxTree.Children);
        var image = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineImage);

        Assert.Collection(image.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageAlt, node.Kind);
                Assert.Equal("foo *bar*", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 16), node.SourceSpan);
            },
            node => Assert.Equal(MarkdownSyntaxKind.ImageSource, node.Kind),
            node => Assert.Equal(MarkdownSyntaxKind.ImageTitle, node.Kind));

        var html = result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });

        Assert.Equal("<p>Lead <img src=\"train.jpg\" alt=\"foo bar\" title=\"train &amp; tracks\" /></p>", html);
    }

    [Fact]
    public void ParseWithSyntaxTree_CommonMark_Profile_Leaves_Standalone_Image_Lines_Inside_Paragraphs() {
        const string markdown = "![foo *bar*](train.jpg \"train & tracks\")\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var paragraph = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);

        var image = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.InlineImage);
        Assert.Equal("train.jpg", image.Literal);
        Assert.Collection(image.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageAlt, node.Kind);
                Assert.Equal("foo *bar*", node.Literal);
            },
            node => Assert.Equal(MarkdownSyntaxKind.ImageSource, node.Kind),
            node => Assert.Equal(MarkdownSyntaxKind.ImageTitle, node.Kind));
    }

    [Fact]
    public void ParseWithSyntaxTree_CommonMark_Profile_Uses_Full_PostMarker_Padding_For_List_Continuation_Spans() {
        const string markdown = """
 -    one

      two
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var list = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);

        var item = Assert.Single(list.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, item.Kind);
        Assert.Equal(2, item.Children.Count);

        var lead = item.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, lead.Kind);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 9), lead.SourceSpan);
        Assert.Equal("one", lead.Literal);

        var trailing = item.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, trailing.Kind);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 9), trailing.SourceSpan);
        Assert.Equal("two", trailing.Literal);

        var deepest = result.FindDeepestNodeAtPosition(3, 8);
        Assert.NotNull(deepest);
        Assert.Equal(MarkdownSyntaxKind.InlineText, deepest!.Kind);
        Assert.Equal("two", deepest.Literal);

        var nearestBlock = result.FindNearestBlockAtPosition(3, 8);
        Assert.NotNull(nearestBlock);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, nearestBlock!.Kind);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 9), nearestBlock.SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTree_CommonMark_Profile_Represents_CodeFirst_List_Items_As_Block_Children() {
        const string markdown = """
1.     indented code
   paragraph

       more code
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var list = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.OrderedList, list.Kind);

        var item = Assert.Single(list.Children);
        Assert.Equal(MarkdownSyntaxKind.ListItem, item.Kind);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.CodeBlock,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.CodeBlock
        }, item.Children.Select(node => node.Kind).ToArray());

        var firstCode = item.Children[0];
        Assert.Equal("indented code", firstCode.Literal);
        var paragraph = item.Children[1];
        Assert.Equal("paragraph", paragraph.Literal);
        var secondCode = item.Children[2];
        Assert.Equal("more code", secondCode.Literal);

        var semanticList = Assert.IsType<OrderedListBlock>(Assert.Single(result.Document.Blocks));
        var semanticItem = Assert.Single(semanticList.Items);
        Assert.Empty(semanticItem.Content.Nodes);
        Assert.Equal(new[] { typeof(CodeBlock), typeof(ParagraphBlock), typeof(CodeBlock) }, semanticItem.Children.Select(child => child.GetType()).ToArray());
    }

    [Fact]
    public void ParseWithSyntaxTree_CommonMark_Profile_Preserves_Empty_List_Items_Inside_A_List() {
        const string markdown = """
- foo
-
- bar
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var list = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);
        Assert.Equal(3, list.Children.Count);

        var middleItem = list.Children[1];
        Assert.Equal(MarkdownSyntaxKind.ListItem, middleItem.Kind);
        var placeholderParagraph = Assert.Single(middleItem.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, placeholderParagraph.Kind);
        Assert.True(string.IsNullOrEmpty(placeholderParagraph.Literal));

        var semanticList = Assert.IsType<UnorderedListBlock>(Assert.Single(result.Document.Blocks));
        var semanticMiddleItem = semanticList.Items[1];
        Assert.Empty(semanticMiddleItem.Content.Nodes);
        Assert.Empty(semanticMiddleItem.AdditionalParagraphs);
        Assert.Empty(semanticMiddleItem.Children);
    }

    [Fact]
    public void ParseWithSyntaxTree_CommonMark_Profile_Keeps_Shallowly_Indented_Sibling_List_Items_At_The_Same_Level() {
        const string markdown = """
- foo
 - bar
  - baz
   - boo
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var list = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);
        Assert.Equal(4, list.Children.Count);
        Assert.All(list.Children, child => Assert.Equal(MarkdownSyntaxKind.ListItem, child.Kind));

        var semanticList = Assert.IsType<UnorderedListBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal(new[] { 0, 0, 0, 0 }, semanticList.Items.Select(item => item.Level).ToArray());
        Assert.All(semanticList.Items, item => Assert.Empty(item.Children));
    }

    [Fact]
    public void ParseWithSyntaxTree_CommonMark_Profile_Represents_ListFirst_And_HeadingFirst_Items_As_Block_Children() {
        const string markdown = """
- - foo
- # Bar
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var outerList = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, outerList.Kind);
        Assert.Equal(2, outerList.Children.Count);

        var nestedListItem = outerList.Children[0];
        Assert.Equal(MarkdownSyntaxKind.ListItem, nestedListItem.Kind);
        var nestedList = Assert.Single(nestedListItem.Children);
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, nestedList.Kind);

        var headingItem = outerList.Children[1];
        Assert.Equal(MarkdownSyntaxKind.ListItem, headingItem.Kind);
        var heading = Assert.Single(headingItem.Children);
        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.Equal("Bar", heading.Literal);

        var semanticOuterList = Assert.IsType<UnorderedListBlock>(Assert.Single(result.Document.Blocks));
        Assert.IsType<UnorderedListBlock>(semanticOuterList.Items[0].Children.Single());
        Assert.IsType<HeadingBlock>(semanticOuterList.Items[1].Children.Single());
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Reference_Link_Metadata_From_Definitions() {
        var markdown = """
[Full][hero] [collapsed][] [shortcut]

[hero]: https://example.com/full "Full title"
[collapsed]: https://example.com/collapsed "Collapsed title"
[shortcut]: https://example.com/shortcut "Shortcut title"
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Equal(new[] {
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.ReferenceLinkDefinition,
            MarkdownSyntaxKind.ReferenceLinkDefinition,
            MarkdownSyntaxKind.ReferenceLinkDefinition
        }, result.SyntaxTree.Children.Select(node => node.Kind).ToArray());

        Assert.Single(result.FinalSyntaxTree.Children);

        var paragraph = result.SyntaxTree.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineLink,
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineLink,
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineLink
        }, paragraph.Children.Select(node => node.Kind).ToArray());

        var full = paragraph.Children[0];
        Assert.Collection(full.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineText, node.Kind);
                Assert.Equal("Full", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, node.Kind);
                Assert.Equal("https://example.com/full", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(3, 9, 3, 32), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTitle, node.Kind);
                Assert.Equal("Full title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(3, 35, 3, 44), node.SourceSpan);
            });

        var collapsed = paragraph.Children[2];
        Assert.Collection(collapsed.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineText, node.Kind);
                Assert.Equal("collapsed", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, node.Kind);
                Assert.Equal("https://example.com/collapsed", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(4, 14, 4, 42), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTitle, node.Kind);
                Assert.Equal("Collapsed title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(4, 45, 4, 59), node.SourceSpan);
            });

        var shortcut = paragraph.Children[4];
        Assert.Collection(shortcut.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineText, node.Kind);
                Assert.Equal("shortcut", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, node.Kind);
                Assert.Equal("https://example.com/shortcut", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(5, 13, 5, 40), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTitle, node.Kind);
                Assert.Equal("Shortcut title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(5, 43, 5, 56), node.SourceSpan);
            });
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Reference_Image_Metadata_From_Definitions() {
        var markdown = """
See ![Badge][hero]

[hero]: https://example.com/badge.svg "Build badge"
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Equal(new[] {
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.ReferenceLinkDefinition
        }, result.SyntaxTree.Children.Select(node => node.Kind).ToArray());

        Assert.Single(result.FinalSyntaxTree.Children);

        var paragraph = result.SyntaxTree.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineImage
        }, paragraph.Children.Select(node => node.Kind).ToArray());

        var image = paragraph.Children[1];
        Assert.Equal("https://example.com/badge.svg", image.Literal);
        Assert.Collection(image.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageAlt, node.Kind);
                Assert.Equal("Badge", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 11), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageSource, node.Kind);
                Assert.Equal("https://example.com/badge.svg", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(3, 9, 3, 37), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageTitle, node.Kind);
                Assert.Equal("Build badge", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(3, 40, 3, 50), node.SourceSpan);
            });
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Reference_Definition_Syntax_Nodes_And_Position_Lookups() {
        var markdown = """
[hero]

[hero]: https://example.com/docs
  "Docs title"
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Equal(new[] {
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.ReferenceLinkDefinition
        }, result.SyntaxTree.Children.Select(node => node.Kind).ToArray());

        var definition = result.SyntaxTree.Children[1];
        Assert.Equal(new MarkdownSourceSpan(3, 1, 4, 14), definition.SourceSpan);
        Assert.Collection(definition.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ReferenceLinkLabel, node.Kind);
                Assert.Equal("hero", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(3, 2, 3, 5), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ReferenceLinkUrl, node.Kind);
                Assert.Equal("https://example.com/docs", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(3, 9, 3, 32), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ReferenceLinkTitle, node.Kind);
                Assert.Equal("Docs title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(4, 4, 4, 13), node.SourceSpan);
            });

        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkUrl, result.FindDeepestNodeAtPosition(3, 15)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkTitle, result.FindDeepestNodeAtPosition(4, 6)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkDefinition, result.FindNearestBlockAtPosition(4, 6)!.Kind);
        Assert.Null(result.FindDeepestFinalNodeAtPosition(3, 15));
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Multiline_Reference_Definition_Destination_And_Title() {
        var markdown = """
[Foo bar]

[Foo bar]:
<my url>
'title'
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        Assert.Equal(new[] {
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.ReferenceLinkDefinition
        }, result.SyntaxTree.Children.Select(node => node.Kind).ToArray());

        var paragraph = result.SyntaxTree.Children[0];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        var link = Assert.Single(paragraph.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineLink, link.Kind);
        Assert.Collection(link.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineText, node.Kind);
                Assert.Equal("Foo bar", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, node.Kind);
                Assert.Equal("my url", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(4, 2, 4, 7), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTitle, node.Kind);
                Assert.Equal("title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(5, 2, 5, 6), node.SourceSpan);
            });

        var definition = result.SyntaxTree.Children[1];
        Assert.Equal(new MarkdownSourceSpan(3, 1, 5, 7), definition.SourceSpan);
        Assert.Collection(definition.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ReferenceLinkLabel, node.Kind);
                Assert.Equal("foo bar", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(3, 2, 3, 8), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ReferenceLinkUrl, node.Kind);
                Assert.Equal("my url", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(4, 2, 4, 7), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ReferenceLinkTitle, node.Kind);
                Assert.Equal("title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(5, 2, 5, 6), node.SourceSpan);
            });

        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkUrl, result.FindDeepestNodeAtPosition(4, 3)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkTitle, result.FindDeepestNodeAtPosition(5, 3)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkDefinition, result.FindNearestBlockAtPosition(5, 3)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTree_Clears_Definition_Source_Spans_From_Final_Reference_Link_Metadata() {
        var markdown = """
[hero]: https://example.com/docs
  "Docs title"

[hero]
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 6), paragraph.SourceSpan);

        var link = Assert.Single(paragraph.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineLink, link.Kind);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 6), link.SourceSpan);

        Assert.Collection(link.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineText, node.Kind);
                Assert.Equal("hero", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(4, 2, 4, 5), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, node.Kind);
                Assert.Equal("https://example.com/docs", node.Literal);
                Assert.Null(node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTitle, node.Kind);
                Assert.Equal("Docs title", node.Literal);
                Assert.Null(node.SourceSpan);
            });

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Multiline_Reference_Definition_Label_Span() {
        var markdown = """
[Foo
  bar]: /url

[Baz][Foo bar]
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        Assert.Equal(new[] {
            MarkdownSyntaxKind.ReferenceLinkDefinition,
            MarkdownSyntaxKind.Paragraph
        }, result.SyntaxTree.Children.Select(node => node.Kind).ToArray());

        var definition = result.SyntaxTree.Children[0];
        Assert.Equal(new MarkdownSourceSpan(1, 1, 2, 12), definition.SourceSpan);
        Assert.Collection(definition.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.ReferenceLinkLabel, node.Kind);
                Assert.Equal("foo bar", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 2, 2, 5), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ReferenceLinkUrl, node.Kind);
                Assert.Equal("/url", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(2, 9, 2, 12), node.SourceSpan);
            });

        var paragraph = result.SyntaxTree.Children[1];
        var link = Assert.Single(paragraph.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineLink, link.Kind);
        Assert.Collection(link.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineText, node.Kind);
                Assert.Equal("Baz", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, node.Kind);
                Assert.Equal("/url", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(2, 9, 2, 12), node.SourceSpan);
            });

        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkLabel, result.FindDeepestNodeAtPosition(2, 4)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkDefinition, result.FindNearestBlockAtPosition(2, 4)!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTree_Preserves_Source_Order_When_Reference_Definitions_Precede_Content() {
        var markdown = """
[hero]: https://example.com/docs "Docs title"

[hero]
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Equal(new[] {
            MarkdownSyntaxKind.ReferenceLinkDefinition,
            MarkdownSyntaxKind.Paragraph
        }, result.SyntaxTree.Children.Select(node => node.Kind).ToArray());

        Assert.Equal(MarkdownSyntaxKind.ReferenceLinkDefinition, result.FindNearestBlockAtPosition(1, 10)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, result.FindNearestBlockAtPosition(3, 3)!.Kind);
        Assert.Single(result.FinalSyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, result.FinalSyntaxTree.Children[0].Kind);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 6), result.FinalSyntaxTree.Children[0].SourceSpan);
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
    public void ParseWithSyntaxTree_Assigns_Absolute_SourceSpans_To_Nested_Quote_ObjectModel() {
        var result = MarkdownReader.ParseWithSyntaxTree("""
> quoted
> second
""");

        var quote = Assert.IsType<QuoteBlock>(Assert.Single(result.Document.Blocks));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 8), paragraph.SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTree_Binds_Nested_Quote_Syntax_To_The_Same_Child_Block_Instances() {
        var result = MarkdownReader.ParseWithSyntaxTree("""
> quoted
> second
""");

        var quoteSyntax = Assert.Single(result.SyntaxTree.Children);
        var quoteBlock = Assert.IsType<QuoteBlock>(Assert.Single(result.Document.Blocks));
        var paragraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(quoteBlock.ChildBlocks));
        var paragraphSyntax = Assert.Single(quoteSyntax.Children);

        Assert.Same(paragraphBlock, paragraphSyntax.AssociatedObject);
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

        var finalQuoteBlock = Assert.IsType<QuoteBlock>(Assert.Single(result.Document.Blocks));
        var finalQuoteParagraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(finalQuoteBlock.ChildBlocks));
        var finalQuote = Assert.Single(result.FinalSyntaxTree.Children);
        var finalParagraph = Assert.Single(finalQuote.Children);
        var finalText = Assert.Single(finalParagraph.Children);

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalQuoteBlock, finalQuote.AssociatedObject);
        Assert.Same(finalQuoteParagraphBlock, finalParagraph.AssociatedObject);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 8), finalParagraph.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 8), Assert.IsType<ParagraphBlock>(finalParagraph.AssociatedObject).SourceSpan);
        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Absolute_SourceSpans_On_Rewritten_Nested_Quote_Blocks() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
> original
> second
""", options);

        var quote = Assert.IsType<QuoteBlock>(Assert.Single(result.Document.Blocks));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 8), paragraph.SourceSpan);
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
        Assert.Equal(3, callout.Children.Count);

        var kind = callout.Children[0];
        Assert.Equal(MarkdownSyntaxKind.CalloutKind, kind.Kind);
        Assert.Equal("note", kind.Literal);
        Assert.NotNull(kind.SourceSpan);
        Assert.Equal(1, kind.SourceSpan!.Value.StartLine);
        Assert.Equal(1, kind.SourceSpan!.Value.EndLine);
        Assert.Equal(5, kind.SourceSpan!.Value.StartColumn);
        Assert.Equal(8, kind.SourceSpan!.Value.EndColumn);

        var title = callout.Children[1];
        Assert.Equal(MarkdownSyntaxKind.CalloutTitle, title.Kind);
        Assert.Equal("Title", title.Literal);
        Assert.Same(Assert.IsType<CalloutBlock>(Assert.Single(result.Document.Blocks)).TitleInlines, title.AssociatedObject);
        Assert.NotNull(title.SourceSpan);
        Assert.Equal(1, title.SourceSpan!.Value.StartLine);
        Assert.Equal(1, title.SourceSpan!.Value.EndLine);
        Assert.Equal(11, title.SourceSpan!.Value.StartColumn);
        Assert.Equal(15, title.SourceSpan!.Value.EndColumn);

        var titleText = Assert.Single(title.Children);
        Assert.Equal(MarkdownSyntaxKind.InlineText, titleText.Kind);
        Assert.NotNull(titleText.SourceSpan);
        Assert.Equal(11, titleText.SourceSpan!.Value.StartColumn);
        Assert.Equal(15, titleText.SourceSpan!.Value.EndColumn);

        var paragraph = callout.Children[2];
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
    public void ParseWithSyntaxTree_Assigns_Absolute_SourceSpans_To_Nested_Callout_ObjectModel() {
        var result = MarkdownReader.ParseWithSyntaxTree("""
> [!NOTE] Title
> body
""");

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(result.Document.Blocks));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks));

        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 6), paragraph.SourceSpan);
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

        var finalCalloutBlock = Assert.IsType<CalloutBlock>(Assert.Single(result.Document.Blocks));
        var finalCalloutParagraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(finalCalloutBlock.ChildBlocks));
        var finalCallout = Assert.Single(result.FinalSyntaxTree.Children);
        Assert.Equal(3, finalCallout.Children.Count);
        var finalKind = finalCallout.Children[0];
        Assert.Equal(MarkdownSyntaxKind.CalloutKind, finalKind.Kind);
        Assert.Equal("note", finalKind.Literal);
        var finalTitle = finalCallout.Children[1];
        Assert.Equal(MarkdownSyntaxKind.CalloutTitle, finalTitle.Kind);
        Assert.Equal("Title", finalTitle.Literal);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalCalloutBlock, finalCallout.AssociatedObject);
        Assert.Same(finalCalloutBlock.TitleInlines, finalTitle.AssociatedObject);
        var finalParagraph = finalCallout.Children[2];
        var finalText = Assert.Single(finalParagraph.Children);

        Assert.Same(finalCalloutParagraphBlock, finalParagraph.AssociatedObject);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 10), finalParagraph.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 10), Assert.IsType<ParagraphBlock>(finalParagraph.AssociatedObject).SourceSpan);
        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Absolute_SourceSpans_On_Rewritten_Nested_Callout_Blocks() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
> [!NOTE] Title
> original
""", options);

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(result.Document.Blocks));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks));

        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 10), paragraph.SourceSpan);
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
        Assert.Equal(MarkdownSyntaxKind.CalloutKind, callout.Children[0].Kind);
        Assert.Equal(MarkdownSyntaxKind.CalloutTitle, callout.Children[1].Kind);
        var list = callout.Children[2];
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
        Assert.Equal(MarkdownSyntaxKind.CalloutKind, callout.Children[0].Kind);
        Assert.Equal("note", callout.Children[0].Literal);
        var title = callout.Children[1];
        Assert.Equal(MarkdownSyntaxKind.CalloutTitle, title.Kind);
        Assert.Equal("Title with **strong** [link](https://example.com)", title.Literal);
        Assert.Same(Assert.IsType<CalloutBlock>(Assert.Single(result.Document.Blocks)).TitleInlines, title.AssociatedObject);
    }

    [Fact]
    public void ParseWithSyntaxTree_Emits_Callout_Kind_Without_Title_Node_For_Untitled_Callouts() {
        var markdown = """
> [!TIP]
> body
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var callout = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.Callout, callout.Kind);
        Assert.Equal("tip", callout.Literal);
        Assert.Equal(2, callout.Children.Count);

        var kind = callout.Children[0];
        Assert.Equal(MarkdownSyntaxKind.CalloutKind, kind.Kind);
        Assert.Equal("tip", kind.Literal);
        Assert.NotNull(kind.SourceSpan);
        Assert.Equal(5, kind.SourceSpan!.Value.StartColumn);
        Assert.Equal(7, kind.SourceSpan!.Value.EndColumn);

        var paragraph = callout.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Equal("body", paragraph.Literal);
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

        var definitionListBlock = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var semanticFirstGroup = definitionListBlock.Groups[0];
        var semanticFirstValue = Assert.Single(semanticFirstGroup.Definitions);

        Assert.Same(definitionListBlock, definitionList.AssociatedObject);
        Assert.Same(semanticFirstGroup, firstGroup.AssociatedObject);
        Assert.Same(semanticFirstValue, firstValue.AssociatedObject);
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
        Assert.Equal("https://example.com", result.FindDeepestNodeAtPosition(1, 25)!.Literal);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.Document,
            MarkdownSyntaxKind.DefinitionList,
            MarkdownSyntaxKind.DefinitionGroup,
            MarkdownSyntaxKind.DefinitionValue,
            MarkdownSyntaxKind.Paragraph,
            MarkdownSyntaxKind.InlineLink,
            MarkdownSyntaxKind.InlineLinkTarget
        }, result.FindNodePathAtPosition(1, 25).Select(node => node.Kind).ToArray());
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

        var finalDefinitionListBlock = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var finalSemanticGroup = finalDefinitionListBlock.Groups[0];
        var finalSemanticValue = Assert.Single(finalSemanticGroup.Definitions);
        var finalSemanticParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(finalSemanticValue.Blocks));

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalDefinitionListBlock, finalDefinitionList.AssociatedObject);
        Assert.Same(finalSemanticGroup, finalFirstGroup.AssociatedObject);
        Assert.Same(finalSemanticValue, finalValue.AssociatedObject);
        Assert.Same(finalSemanticParagraph, finalParagraph.AssociatedObject);
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

        var finalDetailsBlock = Assert.IsType<DetailsBlock>(Assert.Single(result.Document.Blocks));
        var finalDetailsParagraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(finalDetailsBlock.ChildBlocks));
        var finalDetails = Assert.Single(result.FinalSyntaxTree.Children);
        Assert.Equal(2, finalDetails.Children.Count);
        var finalParagraph = finalDetails.Children[1];
        var finalText = Assert.Single(finalParagraph.Children);

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalDetailsBlock, finalDetails.AssociatedObject);
        Assert.Same(finalDetailsParagraphBlock, finalParagraph.AssociatedObject);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, finalParagraph.Kind);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 8), finalParagraph.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 8), Assert.IsType<ParagraphBlock>(finalParagraph.AssociatedObject).SourceSpan);
        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Absolute_SourceSpans_On_Rewritten_Nested_Details_Blocks() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
<details>
<summary>Summary</summary>

original
</details>
""", options);

        var details = Assert.IsType<DetailsBlock>(Assert.Single(result.Document.Blocks));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(details.ChildBlocks));

        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 8), paragraph.SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Transform_Sees_SyntaxTree_And_Nested_Block_SourceSpans() {
        var transform = new CaptureQuoteSyntaxAndBlockSpansTransform();
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(transform);

        _ = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
> original
> second
""", options);

        Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 8), transform.SyntaxTreeParagraphSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 8), transform.BlockParagraphSpan);
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
        Assert.Equal(3, footnote.Children.Count);

        var label = footnote.Children[0];
        Assert.Equal(MarkdownSyntaxKind.FootnoteLabel, label.Kind);
        Assert.Equal("1", label.Literal);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 3), label.SourceSpan);

        var firstParagraph = footnote.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, firstParagraph.Kind);
        Assert.NotNull(firstParagraph.SourceSpan);
        Assert.Equal(3, firstParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(4, firstParagraph.SourceSpan!.Value.EndLine);
        Assert.Equal("first line continued", firstParagraph.Literal);

        var secondParagraph = footnote.Children[2];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, secondParagraph.Kind);
        Assert.NotNull(secondParagraph.SourceSpan);
        Assert.Equal(6, secondParagraph.SourceSpan!.Value.StartLine);
        Assert.Equal(6, secondParagraph.SourceSpan!.Value.EndLine);
        Assert.Equal("second paragraph", secondParagraph.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTree_Assigns_Absolute_SourceSpans_To_Footnote_Paragraph_ObjectModel() {
        var result = MarkdownReader.ParseWithSyntaxTree("""
Lead[^1]

[^1]: first line
""");

        var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(result.Document.Blocks, block => block is FootnoteDefinitionBlock));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(footnote.Blocks));

        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 16), paragraph.SourceSpan);
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
        Assert.Equal(3, footnote.Children.Count);

        var label = footnote.Children[0];
        Assert.Equal(MarkdownSyntaxKind.FootnoteLabel, label.Kind);
        Assert.Equal("1", label.Literal);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 3), label.SourceSpan);

        var intro = footnote.Children[1];
        Assert.Equal(MarkdownSyntaxKind.Paragraph, intro.Kind);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 11), intro.SourceSpan);

        var list = footnote.Children[2];
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
    public void ParseWithSyntaxTree_Assigns_SourceSpan_To_FootnoteDefinition_ObjectModel() {
        var markdown = """
Lead[^1]

[^1]: first line
  continued

  second paragraph
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(result.Document.Blocks, block => block is FootnoteDefinitionBlock));
        Assert.Equal(new MarkdownSourceSpan(3, 1, 6, 18), footnote.SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Rebuilds_Final_Footnote_Syntax_After_Nested_Transform() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
Lead[^1]

[^1]: original
""", options);

        var finalFootnoteBlock = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(result.Document.Blocks, block => block is FootnoteDefinitionBlock));
        var finalFootnoteParagraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(finalFootnoteBlock.Blocks));
        var finalFootnote = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.FootnoteDefinition);
        Assert.Equal(2, finalFootnote.Children.Count);
        var finalLabel = finalFootnote.Children[0];
        Assert.Equal(MarkdownSyntaxKind.FootnoteLabel, finalLabel.Kind);
        Assert.Equal("1", finalLabel.Literal);
        var finalParagraph = finalFootnote.Children[1];
        var finalText = Assert.Single(finalParagraph.Children);

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalFootnoteBlock, finalFootnote.AssociatedObject);
        Assert.Same(finalFootnoteParagraphBlock, finalParagraph.AssociatedObject);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 14), finalParagraph.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 14), Assert.IsType<ParagraphBlock>(finalParagraph.AssociatedObject).SourceSpan);
        Assert.Equal("rewritten", finalParagraph.Literal);
        Assert.Equal("rewritten", finalText.Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Absolute_SourceSpans_On_Rewritten_Nested_Footnote_Blocks() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new RewriteNestedParagraphsTransform("rewritten"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
Lead[^1]

[^1]: original
""", options);

        var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(result.Document.Blocks, block => block is FootnoteDefinitionBlock));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(footnote.Blocks));

        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 14), paragraph.SourceSpan);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Aggregate_SourceSpans_On_Merged_Nested_Quote_Blocks() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsInNestedBlockListsTransform("merged"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
> alpha
>
> beta
""", options);

        var quote = Assert.IsType<QuoteBlock>(Assert.Single(result.Document.Blocks));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        Assert.Equal(new MarkdownSourceSpan(1, 3, 3, 6), paragraph.SourceSpan);

        var finalQuote = Assert.Single(result.FinalSyntaxTree.Children);
        var finalParagraph = Assert.Single(finalQuote.Children);
        var finalText = Assert.Single(finalParagraph.Children);

        Assert.Equal(new MarkdownSourceSpan(1, 3, 3, 6), finalParagraph.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 3, 6), Assert.IsType<ParagraphBlock>(finalParagraph.AssociatedObject).SourceSpan);
        Assert.Equal("merged", finalParagraph.Literal);
        Assert.Equal("merged", finalText.Literal);
        Assert.Null(finalText.SourceSpan);

        var diagnostic = Assert.Single(result.TransformDiagnostics);
        Assert.Equal("Document > Quote", diagnostic.AffectedOriginalBlockPath);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 6), diagnostic.AffectedOriginalBlockSpan);
        Assert.Equal("Document > Quote", diagnostic.AffectedFinalBlockPath);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 6), diagnostic.AffectedFinalBlockSpan);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Final_Quote_List_Syntax_Associations_For_MultiParagraph_List_Items() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsInNestedBlockListsTransform("merged"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
> - alpha
>
>   beta
""", options);

        var finalQuoteBlock = Assert.IsType<QuoteBlock>(Assert.Single(result.Document.Blocks));
        var finalListBlock = Assert.IsType<UnorderedListBlock>(Assert.Single(finalQuoteBlock.ChildBlocks));
        var finalListItemBlock = Assert.IsType<ListItem>(Assert.Single(finalListBlock.Items));
        var finalParagraphBlocks = finalListItemBlock.ParagraphBlocks.ToArray();

        var finalQuote = Assert.Single(result.FinalSyntaxTree.Children);
        var finalList = Assert.Single(finalQuote.Children);
        var finalListItem = Assert.Single(finalList.Children);
        var finalParagraphs = finalListItem.Children.ToArray();

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalQuoteBlock, finalQuote.AssociatedObject);
        Assert.Same(finalListBlock, finalList.AssociatedObject);
        Assert.Same(finalListItemBlock, finalListItem.AssociatedObject);
        Assert.Equal(2, finalParagraphBlocks.Length);
        Assert.Equal(2, finalParagraphs.Length);
        Assert.Same(finalParagraphBlocks[0], finalParagraphs[0].AssociatedObject);
        Assert.Same(finalParagraphBlocks[1], finalParagraphs[1].AssociatedObject);
        Assert.Equal("alpha", finalParagraphs[0].Literal);
        Assert.Equal("beta", finalParagraphs[1].Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Final_Callout_List_Syntax_Associations_For_MultiParagraph_List_Items() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsInNestedBlockListsTransform("merged"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
> [!NOTE] Title
> - alpha
>
>   beta
""", options);

        var finalCalloutBlock = Assert.IsType<CalloutBlock>(Assert.Single(result.Document.Blocks));
        var finalListBlock = Assert.IsType<UnorderedListBlock>(Assert.Single(finalCalloutBlock.ChildBlocks));
        var finalListItemBlock = Assert.IsType<ListItem>(Assert.Single(finalListBlock.Items));
        var finalParagraphBlocks = finalListItemBlock.ParagraphBlocks.ToArray();

        var finalCallout = Assert.Single(result.FinalSyntaxTree.Children);
        Assert.Equal(3, finalCallout.Children.Count);
        var finalTitle = finalCallout.Children[1];
        var finalList = finalCallout.Children[2];
        var finalListItem = Assert.Single(finalList.Children);
        var finalParagraphs = finalListItem.Children.ToArray();

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalCalloutBlock, finalCallout.AssociatedObject);
        Assert.Same(finalCalloutBlock.TitleInlines, finalTitle.AssociatedObject);
        Assert.Same(finalListBlock, finalList.AssociatedObject);
        Assert.Same(finalListItemBlock, finalListItem.AssociatedObject);
        Assert.Equal(2, finalParagraphBlocks.Length);
        Assert.Equal(2, finalParagraphs.Length);
        Assert.Same(finalParagraphBlocks[0], finalParagraphs[0].AssociatedObject);
        Assert.Same(finalParagraphBlocks[1], finalParagraphs[1].AssociatedObject);
        Assert.Equal("alpha", finalParagraphs[0].Literal);
        Assert.Equal("beta", finalParagraphs[1].Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Final_Details_List_Syntax_Associations_For_MultiParagraph_List_Items() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsInNestedBlockListsTransform("merged"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
<details>
<summary>Summary</summary>

- alpha

  beta
</details>
""", options);

        var finalDetailsBlock = Assert.IsType<DetailsBlock>(Assert.Single(result.Document.Blocks));
        var finalListBlock = Assert.IsType<UnorderedListBlock>(Assert.Single(finalDetailsBlock.ChildBlocks));
        var finalListItemBlock = Assert.IsType<ListItem>(Assert.Single(finalListBlock.Items));
        var finalParagraphBlocks = finalListItemBlock.ParagraphBlocks.ToArray();

        var finalDetails = Assert.Single(result.FinalSyntaxTree.Children);
        Assert.Equal(2, finalDetails.Children.Count);
        var finalList = finalDetails.Children[1];
        var finalListItem = Assert.Single(finalList.Children);
        var finalParagraphs = finalListItem.Children.ToArray();

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalDetailsBlock, finalDetails.AssociatedObject);
        Assert.Same(finalListBlock, finalList.AssociatedObject);
        Assert.Same(finalListItemBlock, finalListItem.AssociatedObject);
        Assert.Equal(2, finalParagraphBlocks.Length);
        Assert.Equal(2, finalParagraphs.Length);
        Assert.Same(finalParagraphBlocks[0], finalParagraphs[0].AssociatedObject);
        Assert.Same(finalParagraphBlocks[1], finalParagraphs[1].AssociatedObject);
        Assert.Equal("alpha", finalParagraphs[0].Literal);
        Assert.Equal("beta", finalParagraphs[1].Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Final_Footnote_List_Syntax_Associations_For_MultiParagraph_List_Items() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsInNestedBlockListsTransform("merged"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
Lead[^1]

[^1]:
  - alpha

    beta
""", options);

        var finalFootnoteBlock = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(result.Document.Blocks, block => block is FootnoteDefinitionBlock));
        var finalListBlock = Assert.IsType<UnorderedListBlock>(Assert.Single(finalFootnoteBlock.Blocks));
        var finalListItemBlock = Assert.IsType<ListItem>(Assert.Single(finalListBlock.Items));
        var finalParagraphBlocks = finalListItemBlock.ParagraphBlocks.ToArray();

        var finalFootnote = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.FootnoteDefinition);
        Assert.Equal(2, finalFootnote.Children.Count);
        var finalList = finalFootnote.Children[1];
        var finalListItem = Assert.Single(finalList.Children);
        var finalParagraphs = finalListItem.Children.ToArray();

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(finalFootnoteBlock, finalFootnote.AssociatedObject);
        Assert.Same(finalListBlock, finalList.AssociatedObject);
        Assert.Same(finalListItemBlock, finalListItem.AssociatedObject);
        Assert.Equal(2, finalParagraphBlocks.Length);
        Assert.Equal(2, finalParagraphs.Length);
        Assert.Same(finalParagraphBlocks[0], finalParagraphs[0].AssociatedObject);
        Assert.Same(finalParagraphBlocks[1], finalParagraphs[1].AssociatedObject);
        Assert.Equal("alpha", finalParagraphs[0].Literal);
        Assert.Equal("beta", finalParagraphs[1].Literal);
    }

    [Fact]
    public void ParseWithSyntaxTreeAndDiagnostics_Preserves_Aggregate_SourceSpans_On_Split_Nested_Footnote_Blocks() {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new SplitFirstParagraphInNestedBlockListsTransform("first", "second"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
Lead[^1]

[^1]: alpha beta
""", options);

        var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(result.Document.Blocks, block => block is FootnoteDefinitionBlock));
        Assert.Equal(2, footnote.Blocks.Count);
        Assert.All(footnote.Blocks, block => Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 16), Assert.IsType<ParagraphBlock>(block).SourceSpan));

        var finalFootnote = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.FootnoteDefinition);
        Assert.Equal(3, finalFootnote.Children.Count);
        var finalParagraphs = finalFootnote.Children.Skip(1).ToArray();
        Assert.Equal(2, finalParagraphs.Length);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Same(footnote, finalFootnote.AssociatedObject);
        Assert.Same(footnote.Blocks[0], finalParagraphs[0].AssociatedObject);
        Assert.Same(footnote.Blocks[1], finalParagraphs[1].AssociatedObject);
        Assert.All(finalParagraphs, paragraph => Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 16), paragraph.SourceSpan));
        Assert.All(finalParagraphs, paragraph => Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 16), Assert.IsType<ParagraphBlock>(paragraph.AssociatedObject).SourceSpan));
        Assert.Equal("first", finalParagraphs[0].Literal);
        Assert.Equal("second", finalParagraphs[1].Literal);
        Assert.All(finalParagraphs.Select(paragraph => Assert.Single(paragraph.Children)), textNode => Assert.Null(textNode.SourceSpan));
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
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 10), alt.SourceSpan);

        var source = image.Children[1];
        Assert.Equal(MarkdownSyntaxKind.ImageSource, source.Kind);
        Assert.Equal("https://example.com/image.png", source.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 41), source.SourceSpan);

        var title = image.Children[2];
        Assert.Equal(MarkdownSyntaxKind.ImageTitle, title.Kind);
        Assert.Equal("Image title", title.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 44, 1, 54), title.SourceSpan);

        Assert.Equal(MarkdownSyntaxKind.ImageSource, result.FindDeepestNodeAtPosition(1, 20)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.ImageTitle, result.FindDeepestNodeAtPosition(1, 45)!.Kind);
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
                Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 11), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageSource, node.Kind);
                Assert.Equal("https://example.com/image.png", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 14, 1, 42), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTarget, node.Kind);
                Assert.Equal("https://example.com/docs", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 60, 1, 83), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageLinkTitle, node.Kind);
                Assert.Equal("Link title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 86, 1, 95), node.SourceSpan);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.ImageTitle, node.Kind);
                Assert.Equal("Image title", node.Literal);
                Assert.Equal(new MarkdownSourceSpan(1, 45, 1, 55), node.SourceSpan);
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
    public void HtmlImported_Inline_Link_SyntaxNode_Captures_Linked_Html_Metadata() {
        const string html = """
<p><a href="/docs/hero" title="Hero docs" target="_blank" rel="nofollow sponsored">Read more</a></p>
""";

        var document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var syntax = ((ISyntaxMarkdownBlock)paragraph).BuildSyntaxNode(null);
        var link = Assert.Single(syntax.Children);

        Assert.Equal(MarkdownSyntaxKind.InlineLink, link.Kind);
        Assert.Collection(link.Children,
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineText, node.Kind);
                Assert.Equal("Read more", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, node.Kind);
                Assert.Equal("https://example.com/docs/hero", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkTitle, node.Kind);
                Assert.Equal("Hero docs", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkHtmlTarget, node.Kind);
                Assert.Equal("_blank", node.Literal);
            },
            node => {
                Assert.Equal(MarkdownSyntaxKind.InlineLinkHtmlRel, node.Kind);
                Assert.Equal("nofollow sponsored", node.Literal);
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
    public void ParseWithSyntaxTree_Finds_TableCell_As_Nearest_Block_For_Structured_Cell_Span() {
        const string markdown = """
| Section | Notes |
| --- | --- |
| Alpha | Intro<br><br>> Quoted<br><br>- first<br>- second |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var table = Assert.IsType<TableBlock>(Assert.Single(result.Document.Blocks));
        var cell = table.GetCell(0, 1);

        Assert.NotNull(cell);
        Assert.Equal(new MarkdownSourceSpan(3, 11, 3, 58), cell!.SourceSpan);

        var nearestContaining = result.FindNearestBlockContainingSpan(cell.SourceSpan!.Value);
        Assert.NotNull(nearestContaining);
        Assert.Equal(MarkdownSyntaxKind.TableCell, nearestContaining!.Kind);
    }

    [Fact]
    public void ParseWithSyntaxTree_Finds_TableCell_As_Nearest_Block_For_Positions_Between_Structured_Cell_Children() {
        const string markdown = """
| Section | Notes |
| --- | --- |
| Alpha | Intro<br><br>> Quoted<br><br>- first<br>- second |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        Assert.Equal(MarkdownSyntaxKind.TableCell, result.FindNearestBlockAtPosition(3, 16)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.TableCell, result.FindNearestBlockAtPosition(3, 32)!.Kind);
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

    private sealed class AppendParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            document.Add(new ParagraphBlock(new InlineSequence().Text(text)));
            return document;
        }
    }

    private sealed class RewriteFirstTwoParagraphsTransform(string firstText, string secondText) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var rewritten = MarkdownDoc.Create();
            if (document.DocumentHeader != null) {
                rewritten.Add(document.DocumentHeader);
            }

            for (var i = 0; i < document.Blocks.Count; i++) {
                if (i == 0) {
                    rewritten.Add(new ParagraphBlock(new InlineSequence().Text(firstText)));
                } else if (i == 1) {
                    rewritten.Add(new ParagraphBlock(new InlineSequence().Text(secondText)));
                } else {
                    rewritten.Add(document.Blocks[i]);
                }
            }

            return rewritten;
        }
    }

    private sealed class RewriteSecondParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var rewritten = MarkdownDoc.Create();
            if (document.DocumentHeader != null) {
                rewritten.Add(document.DocumentHeader);
            }

            for (var i = 0; i < document.Blocks.Count; i++) {
                if (i == 1) {
                    rewritten.Add(new ParagraphBlock(new InlineSequence().Text(text)));
                } else {
                    rewritten.Add(document.Blocks[i]);
                }
            }

            return rewritten;
        }
    }

    private sealed class MergeFirstTwoParagraphsTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var rewritten = MarkdownDoc.Create();
            if (document.DocumentHeader != null) {
                rewritten.Add(document.DocumentHeader);
            }

            for (var i = 0; i < document.Blocks.Count; i++) {
                if (i == 0) {
                    rewritten.Add(new ParagraphBlock(new InlineSequence().Text(text)));
                } else if (i > 1) {
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

    private sealed class MergeFirstTwoParagraphsInNestedBlockListsTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            MarkdownDocumentBlockListExpander.RewriteDocument(document, context, (blocks, _) => {
                if (blocks.Count >= 2
                    && blocks[0] is ParagraphBlock
                    && blocks[1] is ParagraphBlock) {
                    return new List<IMarkdownBlock> {
                        new ParagraphBlock(new InlineSequence().Text(text))
                    };
                }

                return blocks.ToList();
            });
            return document;
        }
    }

    private sealed class SplitFirstParagraphInNestedBlockListsTransform(string firstText, string secondText) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            MarkdownDocumentBlockListExpander.RewriteDocument(document, context, (blocks, _) => {
                if (blocks.Count == 1 && blocks[0] is ParagraphBlock) {
                    return new List<IMarkdownBlock> {
                        new ParagraphBlock(new InlineSequence().Text(firstText)),
                        new ParagraphBlock(new InlineSequence().Text(secondText))
                    };
                }

                return blocks.ToList();
            });
            return document;
        }
    }

    private sealed class CaptureQuoteSyntaxAndBlockSpansTransform : IMarkdownDocumentTransform {
        public MarkdownSourceSpan? SyntaxTreeParagraphSpan { get; private set; }
        public MarkdownSourceSpan? BlockParagraphSpan { get; private set; }

        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            var quoteSyntax = Assert.Single(context.SyntaxTree!.Children);
            SyntaxTreeParagraphSpan = Assert.Single(quoteSyntax.Children).SourceSpan;

            var quoteBlock = Assert.IsType<QuoteBlock>(Assert.Single(document.Blocks));
            BlockParagraphSpan = Assert.IsType<ParagraphBlock>(Assert.Single(quoteBlock.ChildBlocks)).SourceSpan;
            return document;
        }
    }


    private sealed class CollectingMarkdownVisitor : MarkdownVisitor {
        public List<string> NodeKinds { get; } = new List<string>();

        protected override void DefaultVisit(MarkdownObject node) {
            NodeKinds.Add(node.GetType().Name);
            base.DefaultVisit(node);
        }
    }

    private sealed class ReplaceParagraphRewriter(string text) : MarkdownRewriter {
        protected override IMarkdownBlock RewriteCurrentBlock(IMarkdownBlock block) =>
            block is ParagraphBlock
                ? new ParagraphBlock(new InlineSequence().Text(text))
                : block;
    }

    private sealed class DoubleBraceInline(InlineSequence inlines) : MarkdownInline, IRenderableMarkdownInline, IContextualHtmlMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline, ISyntaxMarkdownInline {
        public InlineSequence Inlines { get; } = inlines ?? new InlineSequence();

        public string RenderMarkdown() => "{{" + Inlines.RenderMarkdown() + "}}";

        public string RenderHtml() => "<span data-inline=\"double-brace\">" + Inlines.RenderHtml() + "</span>";

        string IContextualHtmlMarkdownInline.RenderHtml(HtmlOptions options) =>
            "<span data-inline=\"double-brace\" data-title=\""
            + System.Net.WebUtility.HtmlEncode(options.Title)
            + "\">"
            + Inlines.RenderHtml()
            + "</span>";

        public void AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);

        public MarkdownSyntaxNode BuildSyntaxNode(MarkdownInlineSyntaxBuilderContext context, MarkdownSourceSpan? span) {
            var children = context.BuildChildren(Inlines);
            return new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Unknown,
                span ?? context.GetAggregateSpan(children),
                literal: RenderMarkdown(),
                children: children,
                associatedObject: this,
                customKind: "double-brace");
        }

        InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;
    }

    private static bool TryParseDoubleBraceInline(MarkdownInlineParserContext context, out MarkdownInlineParseResult result) {
        result = default;
        if (context.CurrentChar != '{'
            || context.Position + 1 >= context.Text.Length
            || context.Text[context.Position + 1] != '{') {
            return false;
        }

        var closing = context.Text.IndexOf("}}", context.Position + 2, StringComparison.Ordinal);
        if (closing < 0) {
            return false;
        }

        var innerLength = closing - (context.Position + 2);
        var nested = context.ParseNestedInlines(2, innerLength);
        result = new MarkdownInlineParseResult(new DoubleBraceInline(nested), closing + 2 - context.Position);
        return true;
    }
}
