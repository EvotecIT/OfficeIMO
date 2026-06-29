using OfficeIMO.Markdown;
using MarkdigMarkdown = Markdig.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_DefinitionList_Ast_Tests {
    [Theory]
    [InlineData("Term\n:   Definition\n")]
    [InlineData("Term\n:\tDefinition\n")]
    public void DefinitionList_MarkdigMarkerSyntax_Matches_MarkdigHtml(string markdown) {
        var office = MarkdownReader.Parse(markdown, CreateMarkdigDefinitionListReaderOptions()).ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
    }

    public static TheoryData<string> MarkdigDefinitionListContinuationCases => new() {
        """
Term
:   First paragraph
    continuation
""",
        """
Term
:   First paragraph
lazy continuation
""",
        """
Term
:   First paragraph
Next term
:   Second paragraph
""",
        """
Term
:   First paragraph

    Second paragraph
""",
        """
Term
:   First paragraph

Second paragraph
""",
        """
Term
:   First paragraph

:   Second paragraph
""",
        """
Term
:   First paragraph
---
""",
        """
Term
:   First paragraph
    - item
""",
        """
Term
:   First paragraph
    1. item
""",
        """
Term
:   First paragraph
    > quote
""",
        "Term\n:   \n\n    code",
        """
Term
:   ```
    code
    ```
""",
        """
Term
:   First paragraph
# Heading
""",
        """
Term
:   First paragraph
- sibling item
""",
        """
Term
:   First paragraph
[ref]: https://example.com
""",
        "Term\n:   \n    code\n"
    };

    [Theory]
    [MemberData(nameof(MarkdigDefinitionListContinuationCases))]
    public void DefinitionList_MarkdigContinuationSyntax_Matches_MarkdigHtml(string markdown) {
        var office = MarkdownReader.Parse(markdown, CreateMarkdigDefinitionListReaderOptions()).ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
    }

    [Fact]
    public void DefinitionList_MarkdigMarkerSyntax_Maps_Multiple_Terms_And_Definitions_To_Grouped_Ast() {
        const string markdown = """
Term 1
Term 2
:   First
:   Second
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];

        Assert.Equal(2, group.TermItems.Count);
        Assert.Equal(2, group.Definitions.Count);
        Assert.Equal(4, definitionList.Entries.Count);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 6), group.TermItems[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 6), group.TermItems[1].SourceSpan);
        Assert.Equal(MarkdownSyntaxKind.DefinitionGroup, syntaxGroup.Kind);
        Assert.Same(group, syntaxGroup.AssociatedObject);
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.DefinitionTerm,
                MarkdownSyntaxKind.DefinitionTerm,
                MarkdownSyntaxKind.DefinitionMarker,
                MarkdownSyntaxKind.DefinitionValue,
                MarkdownSyntaxKind.DefinitionMarker,
                MarkdownSyntaxKind.DefinitionValue
            },
            syntaxGroup.Children.Select(child => child.Kind).ToArray());
        Assert.Same(group.TermItems[0], syntaxGroup.Children[0].AssociatedObject);
        Assert.Same(group.TermItems[1], syntaxGroup.Children[1].AssociatedObject);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 1), syntaxGroup.Children[2].SourceSpan);
        Assert.Same(group.Definitions[0], syntaxGroup.Children[3].AssociatedObject);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 1), syntaxGroup.Children[4].SourceSpan);
        Assert.Same(group.Definitions[1], syntaxGroup.Children[5].AssociatedObject);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_BlankSeparatedTerm_BeforeMarker_Starts_New_MarkdigGroup() {
        const string markdown = """
Term 1

Term 2
:   Definition
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[0]);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[1]);
        var group = Assert.Single(definitionList.Groups);
        var syntaxList = result.SyntaxTree.Children[1];
        var syntaxGroup = Assert.Single(syntaxList.Children);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("Term 1", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("Term 2", Assert.Single(group.TermItems).Markdown);
        Assert.Equal("Definition", Assert.IsType<ParagraphBlock>(Assert.Single(Assert.Single(group.Definitions).Blocks)).Inlines.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.DefinitionTerm,
                MarkdownSyntaxKind.DefinitionMarker,
                MarkdownSyntaxKind.DefinitionValue
            },
            syntaxGroup.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 6), Assert.Single(group.TermItems).SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 1), syntaxGroup.Children[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 5, 4, 14), syntaxGroup.Children[2].SourceSpan);
        Assert.Equal(markdown, written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_MarkdigLazyContinuation_Preserves_SoftBreak_In_Definition_Paragraph_Source() {
        const string markdown = """
Term
:   First paragraph
lazy continuation
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraphSyntax = Assert.Single(definitionValue.Children);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph\nlazy continuation", paragraph.Inlines.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.InlineText,
                MarkdownSyntaxKind.InlineSoftBreak,
                MarkdownSyntaxKind.InlineText
            },
            paragraphSyntax.Children.Select(child => child.Kind).ToArray());
        Assert.IsType<SoftBreakInline>(paragraph.Inlines.Nodes[1]);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 17), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 17), paragraphSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 19, 2, 19), paragraphSyntax.Children[1].SourceSpan);
        Assert.Equal("Term\n:   First paragraph\n    lazy continuation", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        var nativeDefinitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(Assert.Single(Assert.Single(nativeDefinitionList.Groups).Definitions).Children));
        Assert.Equal("First paragraph\nlazy continuation", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 17), definitionBody.SourceSpan);
        Assert.Contains(nativeParagraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.SoftBreak);

        Assert.Same(definition, definitionValue.AssociatedObject);
        Assert.Same(paragraph, paragraphSyntax.AssociatedObject);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_MultipleLazyContinuationLines_Preserve_SoftBreaks_And_WriterReparse() {
        const string markdown = """
Term
:   First paragraph
lazy one
lazy two
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var definition = Assert.Single(Assert.Single(definitionList.Groups).Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraphSyntax = Assert.Single(definitionValue.Children);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph\nlazy one\nlazy two", paragraph.Inlines.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.InlineText,
                MarkdownSyntaxKind.InlineSoftBreak,
                MarkdownSyntaxKind.InlineText,
                MarkdownSyntaxKind.InlineSoftBreak,
                MarkdownSyntaxKind.InlineText
            },
            paragraphSyntax.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 8), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 8), paragraphSyntax.SourceSpan);
        Assert.Equal("Term\n:   First paragraph\n    lazy one\n    lazy two", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        var nativeDefinitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(Assert.Single(Assert.Single(nativeDefinitionList.Groups).Definitions).Children));
        Assert.Equal("First paragraph\nlazy one\nlazy two", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 8), definitionBody.SourceSpan);
        Assert.Equal(2, nativeParagraph.InlineRuns.Count(inline => inline.Kind == MarkdownNativeInlineKind.SoftBreak));

        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_MarkdigBlockContinuation_Stays_In_Definition_Body_Source() {
        const string markdown = """
Term
:   First paragraph
# Heading
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var heading = Assert.IsType<HeadingBlock>(definition.Blocks[1]);
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);

        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 9), definitionValue.SourceSpan);
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.Heading
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 2, 19), definitionValue.Children[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 9), definitionValue.Children[1].SourceSpan);
        Assert.Same(definition, definitionValue.AssociatedObject);
        Assert.Same(paragraph, definitionValue.Children[0].AssociatedObject);
        Assert.Same(heading, definitionValue.Children[1].AssociatedObject);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_MarkdigSetextContinuation_Stays_In_Definition_Body_Source() {
        const string markdown = """
Term
:   First paragraph
---
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var heading = Assert.IsType<HeadingBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var headingSyntax = Assert.Single(definitionValue.Children);

        Assert.Equal(2, heading.Level);
        Assert.Equal("First paragraph", heading.Text);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 3), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 3), headingSyntax.SourceSpan);
        Assert.Same(definition, definitionValue.AssociatedObject);
        Assert.Same(heading, headingSyntax.AssociatedObject);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_SetextContinuation_Stops_LazyContinuation_Before_FollowingParagraph() {
        const string markdown = """
Term
:   First paragraph
---
text
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingParagraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[1]);
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var heading = Assert.IsType<HeadingBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var headingSyntax = Assert.Single(definitionValue.Children);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal(2, heading.Level);
        Assert.Equal("First paragraph", heading.Text);
        Assert.Equal("text", trailingParagraph.Inlines.RenderMarkdown());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 3), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 3), headingSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 4), result.SyntaxTree.Children[1].SourceSpan);
        Assert.Equal("Term\n:   \n    ## First paragraph\n\ntext", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, native.Blocks.Count);
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("## First paragraph", definitionBody.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 3), definitionBody.SourceSpan);
        Assert.Contains("Term\n:   Updated\ntext", native.CreateReplaceEdit(definitionBody, "Updated").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_ThematicBreakContinuation_Stops_LazyContinuation_Before_FollowingParagraph() {
        const string markdown = """
Term
:   First paragraph
***
text
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingParagraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[1]);
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        Assert.Equal(2, definition.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var rule = Assert.IsType<HorizontalRuleBlock>(definition.Blocks[1]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("text", trailingParagraph.Inlines.RenderMarkdown());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 3), definitionValue.SourceSpan);
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.HorizontalRule
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 2, 19), definitionValue.Children[0].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 3), definitionValue.Children[1].SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 4), result.SyntaxTree.Children[1].SourceSpan);
        Assert.Equal("***", rule.MarkerText);
        Assert.Equal("Term\n:   First paragraph\n\n    ---\n\ntext", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, native.Blocks.Count);
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("First paragraph\n\n---", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 3), definitionBody.SourceSpan);
        Assert.Contains("Term\n:   Updated\ntext", native.CreateReplaceEdit(definitionBody, "Updated").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_EmptyMarkdigMarkerContinuation_Strips_FirstContinuationIndent_Source() {
        const string markdown = "Term\n:   \n    code\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraphSyntax = Assert.Single(definitionValue.Children);

        Assert.Equal("code", paragraph.Inlines.RenderMarkdown());
        Assert.Equal(new MarkdownSourceSpan(3, 5, 3, 8), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 5, 3, 8), paragraphSyntax.SourceSpan);
        Assert.Equal("code", definitionValue.Literal);
        Assert.Same(definition, definitionValue.AssociatedObject);
        Assert.Same(paragraph, paragraphSyntax.AssociatedObject);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_EmptyMarkdigMarker_BlankSeparatedBody_Writes_LooseMarkerSyntax_ForReparse() {
        const string markdown = "Term\n:   \n\n    code";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var marker = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionMarker);
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraphSyntax = Assert.Single(definitionValue.Children);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("code", paragraph.Inlines.RenderMarkdown());
        Assert.True(definition.ForceParagraphHtml);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 1), marker.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 5, 4, 8), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 5, 4, 8), paragraphSyntax.SourceSpan);
        Assert.Equal(markdown, written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("code", definitionBody.Value);
        Assert.Equal(new MarkdownSourceSpan(4, 5, 4, 8), definitionBody.SourceSpan);
        Assert.Contains("    updated", native.CreateReplaceEdit(definitionBody, "updated").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_TableShapedContinuation_Stays_Literal_When_Tables_Are_Off() {
        const string markdown = """
Term
:   | A |
    |---|
    | B |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraphSyntax = Assert.Single(definitionValue.Children);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("| A | |---| | B |", NormalizePlainText(InlinePlainText.Extract(paragraph.Inlines)));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 9), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 9), paragraphSyntax.SourceSpan);
        Assert.Contains(@"\| A \|", written, StringComparison.Ordinal);
        Assert.Contains(@"\|---\|", written, StringComparison.Ordinal);
        Assert.Contains(@"\| B \|", written, StringComparison.Ordinal);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_UnindentedTableShapedLazyContinuation_Stays_Literal_When_Tables_Are_Off() {
        const string markdown = """
Term
:   First paragraph
| A |
|---|
| B |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraphSyntax = Assert.Single(definitionValue.Children);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph | A | |---| | B |", NormalizePlainText(InlinePlainText.Extract(paragraph.Inlines)));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 5), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 5), paragraphSyntax.SourceSpan);
        Assert.Contains(@"\| A \|", written, StringComparison.Ordinal);
        Assert.Contains(@"\|---\|", written, StringComparison.Ordinal);
        Assert.Contains(@"\| B \|", written, StringComparison.Ordinal);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("First paragraph\n\\| A \\|\n\\|---\\|\n\\| B \\|", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 5), definitionBody.SourceSpan);
        Assert.Contains(":   updated", native.CreateReplaceEdit(definitionBody, "updated").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_TableShapedContinuation_Becomes_Nested_Table_When_Tables_Are_On() {
        const string markdown = """
Term
:   | A |
    |---|
    | B |
""";

        var readerOptions = CreateMarkdigDefinitionListReaderOptions();
        readerOptions.Tables = true;
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, readerOptions);
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var table = Assert.IsType<TableBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var tableSyntax = Assert.Single(definitionValue.Children);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, readerOptions);
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListAndPipeTablesPipeline());

        Assert.Equal("A", Assert.Single(table.Headers));
        Assert.Equal("B", Assert.Single(Assert.Single(table.Rows)));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 9), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 9), tableSyntax.SourceSpan);
        Assert.Equal("Term\n:   \n    | A |\n    | --- |\n    | B |", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [InlineData("Term\n: Definition\n")]
    [InlineData("Term\n:  Definition\n")]
    public void DefinitionList_MarkdigMarkerSyntax_Requires_MarkdigMarkerSpacing(string markdown) {
        var doc = MarkdownReader.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var html = doc.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
        Assert.DoesNotContain("<dl>", html, StringComparison.Ordinal);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(html));
    }

    [Fact]
    public void DefinitionList_EntryTermMutation_Updates_Grouped_Ast_And_Renderers() {
        var definitionList = new DefinitionListBlock();
        var group = new DefinitionListGroup(
            new[] { new InlineSequence().Text("Original") },
            new[] {
                new DefinitionListDefinition(new[] { new ParagraphBlock(new InlineSequence().Text("first")) }),
                new DefinitionListDefinition(new[] { new ParagraphBlock(new InlineSequence().Text("second")) })
            });
        definitionList.AddGroup(group);

        definitionList.Entries[0].Term = new InlineSequence().Text("Renamed");

        Assert.Equal("Renamed", group.Terms[0].RenderMarkdown());
        Assert.Equal("Renamed", group.TermItems[0].Markdown);
        Assert.Same(group.TermItems[0].Inlines, group.Terms[0]);
        Assert.All(definitionList.Entries, entry => Assert.Equal("Renamed", entry.TermMarkdown));
        Assert.Equal("Renamed\n:   first\n:   second", ((IMarkdownBlock)definitionList).RenderMarkdown());

        var html = ((IMarkdownBlock)definitionList).RenderHtml();
        Assert.Contains("<dt>Renamed</dt>", html);
        Assert.DoesNotContain("Original", html);
    }

    [Fact]
    public void DefinitionList_ProgrammaticGroupedSyntax_Emits_Generated_Marker_Tokens() {
        var definitionList = new DefinitionListBlock();
        definitionList.AddGroup(new DefinitionListGroup(
            new[] {
                new InlineSequence().Text("Term 1"),
                new InlineSequence().Text("Term 2")
            },
            new[] {
                new DefinitionListDefinition(new[] { new ParagraphBlock(new InlineSequence().Text("first")) }),
                new DefinitionListDefinition(new[] { new ParagraphBlock(new InlineSequence().Text("second")) })
            }));
        var document = MarkdownDoc.Create().Add(definitionList);

        var syntaxTree = MarkdownReader.BuildSyntaxTree(document);
        var syntaxGroup = Assert.Single(Assert.Single(syntaxTree.Children).Children);
        var markers = syntaxGroup.Children
            .Where(child => child.Kind == MarkdownSyntaxKind.DefinitionMarker)
            .ToArray();

        Assert.Equal("Term 1\nTerm 2\n:   first\n:   second", NormalizeMarkdown(document.ToMarkdown()));
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.DefinitionTerm,
                MarkdownSyntaxKind.DefinitionTerm,
                MarkdownSyntaxKind.DefinitionMarker,
                MarkdownSyntaxKind.DefinitionValue,
                MarkdownSyntaxKind.DefinitionMarker,
                MarkdownSyntaxKind.DefinitionValue
            },
            syntaxGroup.Children.Select(child => child.Kind).ToArray());
        Assert.All(markers, marker => {
            Assert.Equal(":", marker.Literal);
            Assert.Null(marker.SourceSpan);
        });
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(syntaxTree);
    }

    [Fact]
    public void DefinitionList_MarkdigMarkerSyntax_Writes_Marker_Syntax_For_Reparse() {
        const string markdown = """
Term
:   Definition
""";

        var document = MarkdownReader.Parse(markdown);
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));
        var written = NormalizeMarkdown(document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written);

        Assert.Equal(markdown, written);
        Assert.Equal(markdown, ((IMarkdownBlock)definitionList).RenderMarkdown());
        var reparsedDefinitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(reparsed.Blocks));
        var group = Assert.Single(reparsedDefinitionList.Groups);
        Assert.Equal("Term", Assert.Single(group.TermItems).Markdown);
        Assert.Equal("Definition", Assert.IsType<ParagraphBlock>(Assert.Single(Assert.Single(group.Definitions).Blocks)).Inlines.RenderMarkdown());
    }

    [Fact]
    public void DefinitionList_GroupedTermsAndDefinitions_Write_Without_Flattening() {
        const string markdown = """
Term 1
Term 2
:   First
:   Second
""";

        var document = MarkdownReader.Parse(markdown);
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));
        var written = NormalizeMarkdown(document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written);

        Assert.Equal(markdown, written);
        Assert.Equal(markdown, ((IMarkdownBlock)definitionList).RenderMarkdown());
        var reparsedDefinitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(reparsed.Blocks));
        var group = Assert.Single(reparsedDefinitionList.Groups);
        Assert.Equal(new[] { "Term 1", "Term 2" }, group.TermItems.Select(term => term.Markdown).ToArray());
        Assert.Equal(
            new[] { "First", "Second" },
            group.Definitions
                .Select(definition => Assert.IsType<ParagraphBlock>(Assert.Single(definition.Blocks)).Inlines.RenderMarkdown())
                .ToArray());
        Assert.Equal(4, definitionList.Entries.Count);
    }

    [Fact]
    public void DefinitionList_BlankSeparatedDefinitions_Write_LooseMarkerSyntax_ForReparse() {
        const string markdown = """
Term
:   First paragraph

:   Second paragraph
""";

        var document = MarkdownReader.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var written = NormalizeMarkdown(document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal(markdown, written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
    }

    [Fact]
    public void DefinitionList_BlankSeparatedMarkerGroups_Write_BlankSeparator_ForReparse() {
        const string markdown = """
Term 1
:   First

Term 2
:   Second
""";

        var document = MarkdownReader.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var written = NormalizeMarkdown(document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var reparsedDefinitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(reparsed.Blocks));
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal(markdown, written);
        Assert.Equal(2, reparsedDefinitionList.Groups.Count);
        Assert.Equal(new[] { "Term 1", "Term 2" }, reparsedDefinitionList.Groups.Select(group => Assert.Single(group.TermItems).Markdown).ToArray());
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
    }

    [Fact]
    public void DefinitionList_NestedListBody_Writes_MarkerSyntax_ForMarkdigReparse() {
        const string markdown = """
Term
:   First paragraph
    - item
""";

        var document = MarkdownReader.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var written = NormalizeMarkdown(document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal(markdown, written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
    }

    [Fact]
    public void DefinitionList_NestedListBody_Merges_UnindentedSameMarkerLazyListItem() {
        const string markdown = """
Term
:   First paragraph
    - item
- sibling item
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var list = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var listSyntax = Assert.IsType<MarkdownSyntaxNode>(definitionValue.Children[1]);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", paragraph.Inlines.RenderMarkdown());
        Assert.Equal(new[] { "item", "sibling item" }, list.Items.Select(item => item.Content.RenderMarkdown()).ToArray());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.UnorderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(MarkdownSyntaxKind.UnorderedList, listSyntax.Kind);
        Assert.Equal(2, listSyntax.Children.Count(child => child.Kind == MarkdownSyntaxKind.ListItem));
        Assert.Equal(new MarkdownSourceSpan(3, 5, 3, 5), list.Items[0].MarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 1), list.Items[1].MarkerSourceSpan);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedOrderedListBody_Merges_UnindentedLazyOrderedItem() {
        const string markdown = """
Term
:   First paragraph
    1. item
2. sibling item
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var list = Assert.IsType<OrderedListBlock>(definition.Blocks[1]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var listSyntax = definitionValue.Children[1];
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", paragraph.Inlines.RenderMarkdown());
        Assert.Equal(new[] { "item", "sibling item" }, list.Items.Select(item => item.Content.RenderMarkdown()).ToArray());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.OrderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(MarkdownSyntaxKind.OrderedList, listSyntax.Kind);
        Assert.Equal(2, listSyntax.Children.Count(child => child.Kind == MarkdownSyntaxKind.ListItem));
        Assert.Equal(new MarkdownSourceSpan(3, 5, 3, 6), list.Items[0].MarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 2), list.Items[1].MarkerSourceSpan);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedListBody_Keeps_DifferentUnorderedMarkers_Separate() {
        const string markdown = """
Term
:   First paragraph
    - item
* sibling item
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var firstList = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var secondList = Assert.IsType<UnorderedListBlock>(definition.Blocks[2]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("item", Assert.Single(firstList.Items).Content.RenderMarkdown());
        Assert.Equal("sibling item", Assert.Single(secondList.Items).Content.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.UnorderedList,
                MarkdownSyntaxKind.UnorderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_HtmlBlockBody_LazyContinuation_MatchesMarkdigHtml_AndWriterReparse() {
        const string markdown = """
Term
:   <div>
    html
    </div>
lazy
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var html = Assert.IsType<HtmlRawBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("<div>\n  html\n  </div>\nlazy", html.Html.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 4), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 4), Assert.Single(definitionValue.Children).SourceSpan);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("<div>\n  html\n  </div>\nlazy", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 5, 4), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_EmptyMarkerHtmlBlockBody_LazyContinuation_MatchesMarkdigHtml_AndWriterReparse() {
        const string markdown = "Term\n:   \n    <div>\n    html\n    </div>\nlazy\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var html = Assert.IsType<HtmlRawBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("<div>\n  html\n  </div>\nlazy", html.Html.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(3, 5, 6, 4), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 5, 6, 4), Assert.Single(definitionValue.Children).SourceSpan);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("<div>\n  html\n  </div>\nlazy", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(3, 5, 6, 4), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_FencedCodeBody_StopsLazyContinuationAfterClosedFence_AndWriterReparses() {
        const string markdown = """
Term
:   ~~~
    code
    ~~~
lazy
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingParagraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[1]);
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var codeBlock = Assert.IsType<CodeBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("lazy", trailingParagraph.Inlines.RenderMarkdown());
        Assert.Equal("code", codeBlock.Content);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 7), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 7), Assert.Single(definitionValue.Children).SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 4), result.SyntaxTree.Children[1].SourceSpan);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("```\ncode\n```", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 4, 7), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_EmptyMarkerFencedCodeBody_StopsLazyContinuationAfterClosedFence_AndWriterReparses() {
        const string markdown = "Term\n:   \n    ~~~\n    code\n    ~~~\nlazy\n";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingParagraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[1]);
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var codeBlock = Assert.IsType<CodeBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("lazy", trailingParagraph.Inlines.RenderMarkdown());
        Assert.Equal("code", codeBlock.Content);
        Assert.Equal(new MarkdownSourceSpan(3, 5, 5, 7), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 5, 5, 7), Assert.Single(definitionValue.Children).SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(6, 1, 6, 4), result.SyntaxTree.Children[1].SourceSpan);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("```\ncode\n```", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(3, 5, 5, 7), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedListBody_Merges_TableShapedLazyParagraphIntoLastItem() {
        const string markdown = """
Term
:   First
    - item
| A |
|---|
| B |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        Assert.Equal(2, definition.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var nestedList = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var item = Assert.Single(nestedList.Items);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var listSyntax = definitionValue.Children.Single(child => child.Kind == MarkdownSyntaxKind.UnorderedList);
        var listItemSyntax = listSyntax.Children.Single(child => child.Kind == MarkdownSyntaxKind.ListItem);
        var itemParagraphSyntax = listItemSyntax.Children.Single(child => child.Kind == MarkdownSyntaxKind.Paragraph);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("item\n\\| A \\|\n\\|---\\|\n\\| B \\|", item.Content.RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.UnorderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 6, 5), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 6, 5), listSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 5, 6, 5), listItemSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 6, 5), itemParagraphSyntax.SourceSpan);
        Assert.Equal("Term\n:   First\n    - item\n| A |\n|---|\n| B |", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("First\n\n- item\n  \\| A \\|\n  \\|---\\|\n  \\| B \\|", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 6, 5), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedListBody_Stops_Before_UnindentedHeading() {
        const string markdown = """
Term
:   First paragraph
    - item
# Heading
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingHeading = Assert.IsType<HeadingBlock>(result.Document.Blocks[1]);
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        Assert.Equal(2, definition.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var nestedList = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("item", Assert.Single(nestedList.Items).Content.RenderMarkdown());
        Assert.Equal("Heading", trailingHeading.Text);
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.UnorderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 10), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 9), result.SyntaxTree.Children[1].SourceSpan);
        Assert.Contains("\n\n# Heading", written, StringComparison.Ordinal);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, native.Blocks.Count);
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("First paragraph\n\n- item", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 10), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedListBody_Stops_Before_UnindentedThematicBreak() {
        const string markdown = """
Term
:   First paragraph
    - item
***
text
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(3, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingRule = Assert.IsType<HorizontalRuleBlock>(result.Document.Blocks[1]);
        var trailingParagraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[2]);
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        Assert.Equal(2, definition.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var nestedList = Assert.IsType<UnorderedListBlock>(definition.Blocks[1]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("item", Assert.Single(nestedList.Items).Content.RenderMarkdown());
        Assert.Equal("***", trailingRule.MarkerText);
        Assert.Equal("text", trailingParagraph.Inlines.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.UnorderedList
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 10), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 3), result.SyntaxTree.Children[1].SourceSpan);
        Assert.Contains("\n\n---\n\ntext", written, StringComparison.Ordinal);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedBlockquoteBody_Stops_Before_UnindentedHeading() {
        const string markdown = """
Term
:   First paragraph
    > quote
# Heading
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingHeading = Assert.IsType<HeadingBlock>(result.Document.Blocks[1]);
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        Assert.Equal(2, definition.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var quote = Assert.IsType<QuoteBlock>(definition.Blocks[1]);
        var quoteParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("quote", quoteParagraph.Inlines.RenderMarkdown());
        Assert.Equal("Heading", trailingHeading.Text);
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.Quote
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 11), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 9), result.SyntaxTree.Children[1].SourceSpan);
        Assert.Contains("\n\n# Heading", written, StringComparison.Ordinal);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(2, native.Blocks.Count);
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("First paragraph\n\n> quote", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 11), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedBlockquoteBody_Stops_Before_UnindentedThematicBreak() {
        const string markdown = """
Term
:   First paragraph
    > quote
***
text
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        Assert.Equal(3, result.Document.Blocks.Count);
        var definitionList = Assert.IsType<DefinitionListBlock>(result.Document.Blocks[0]);
        var trailingRule = Assert.IsType<HorizontalRuleBlock>(result.Document.Blocks[1]);
        var trailingParagraph = Assert.IsType<ParagraphBlock>(result.Document.Blocks[2]);
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        Assert.Equal(2, definition.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var quote = Assert.IsType<QuoteBlock>(definition.Blocks[1]);
        var quoteParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First paragraph", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("quote", quoteParagraph.Inlines.RenderMarkdown());
        Assert.Equal("***", trailingRule.MarkerText);
        Assert.Equal("text", trailingParagraph.Inlines.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.Quote
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 11), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 3), result.SyntaxTree.Children[1].SourceSpan);
        Assert.Contains("\n\n---\n\ntext", written, StringComparison.Ordinal);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedBlockquoteBody_Merges_TableShapedLazyLinesIntoQuote_WhenTablesAreOff() {
        const string markdown = """
Term
:   First
    > quote
| A |
|---|
| B |
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        Assert.Equal(2, definition.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var quote = Assert.IsType<QuoteBlock>(definition.Blocks[1]);
        var quoteParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var quoteSyntax = definitionValue.Children.Single(child => child.Kind == MarkdownSyntaxKind.Quote);
        var quoteParagraphSyntax = quoteSyntax.Children.Single(child => child.Kind == MarkdownSyntaxKind.Paragraph);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("First", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("quote \\| A \\| \\|---\\| \\| B \\|", quoteParagraph.Inlines.RenderMarkdown());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.Quote
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 6, 5), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 6, 5), quoteSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 6, 5), quoteParagraphSyntax.SourceSpan);
        Assert.Equal("Term\n:   First\n    > quote \\| A \\| \\|---\\| \\| B \\|", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("First\n\n> quote \\| A \\| \\|---\\| \\| B \\|", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 6, 5), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_NestedBlockquoteBody_Merges_TableShapedLazyLinesIntoQuote_WhenTablesAreOn() {
        const string markdown = """
Term
:   First
    > quote
| A |
|---|
| B |
""";

        var readerOptions = CreateMarkdigDefinitionListReaderOptions();
        readerOptions.Tables = true;
        var result = MarkdownReader.ParseWithSyntaxTree(markdown, readerOptions);
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        Assert.Equal(2, definition.Blocks.Count);
        var paragraph = Assert.IsType<ParagraphBlock>(definition.Blocks[0]);
        var quote = Assert.IsType<QuoteBlock>(definition.Blocks[1]);
        Assert.Equal(2, quote.ChildBlocks.Count);
        var quoteParagraph = Assert.IsType<ParagraphBlock>(quote.ChildBlocks[0]);
        var table = Assert.IsType<TableBlock>(quote.ChildBlocks[1]);
        var syntaxGroup = result.SyntaxTree.Children[0].Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var quoteSyntax = definitionValue.Children.Single(child => child.Kind == MarkdownSyntaxKind.Quote);
        var quoteParagraphSyntax = quoteSyntax.Children.Single(child => child.Kind == MarkdownSyntaxKind.Paragraph);
        var tableSyntax = quoteSyntax.Children.Single(child => child.Kind == MarkdownSyntaxKind.Table);
        var written = NormalizeMarkdown(result.Document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, readerOptions);
        var office = result.Document.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var reparsedOffice = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListAndPipeTablesPipeline());

        Assert.Equal("First", paragraph.Inlines.RenderMarkdown());
        Assert.Equal("quote", quoteParagraph.Inlines.RenderMarkdown());
        Assert.Equal("A", Assert.Single(table.Headers));
        Assert.Equal("B", Assert.Single(Assert.Single(table.Rows)));
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.Quote
            },
            definitionValue.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(
            new[] {
                MarkdownSyntaxKind.QuoteMarker,
                MarkdownSyntaxKind.Paragraph,
                MarkdownSyntaxKind.Table
            },
            quoteSyntax.Children.Select(child => child.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 6, 5), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 6, 5), quoteSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 7, 3, 11), quoteParagraphSyntax.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 6, 5), tableSyntax.SourceSpan);
        Assert.Equal("Term\n:   First\n    > quote\n    > \n    > | A |\n    > | --- |\n    > | B |", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

        var native = MarkdownNativeDocument.Parse(markdown, readerOptions);
        var definitionBody = Assert.Single(native.EnumerateBlockSourceFields("definitionBody"));
        Assert.Equal("First\n\n> quote\n> \n> | A |\n> | --- |\n> | B |", definitionBody.Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(2, 5, 6, 5), definitionBody.SourceSpan);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_SetextContinuation_Writes_MarkerSyntax_ForMarkdigReparse() {
        const string markdown = """
Term
:   First paragraph
---
""";

        var document = MarkdownReader.Parse(markdown, CreateMarkdigDefinitionListReaderOptions());
        var written = NormalizeMarkdown(document.ToMarkdown());
        var reparsed = MarkdownReader.Parse(written, CreateMarkdigDefinitionListReaderOptions());
        var office = reparsed.ToHtmlFragment(CreateMarkdigDefinitionListHtmlOptions());
        var markdig = MarkdigMarkdown.ToHtml(markdown, CreateMarkdigDefinitionListPipeline());

        Assert.Equal("Term\n:   \n    ## First paragraph", written);
        Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
    }

    [Fact]
    public void DefinitionList_SimpleInlineSyntax_Still_Writes_Simple_Inline_Syntax() {
        const string markdown = "Term: Definition";

        var document = MarkdownReader.Parse(markdown);
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Equal(markdown, ((IMarkdownBlock)definitionList).RenderMarkdown());
        Assert.Equal(markdown, NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void DefinitionList_EntryTermMutation_Clears_Parsed_SyntaxCache() {
        var result = MarkdownReader.ParseWithSyntaxTree("Original: first");
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));

        Assert.NotEmpty(definitionList.SyntaxItems);

        definitionList.Entries[0].Term = new InlineSequence().Text("Renamed");

        Assert.Empty(definitionList.SyntaxItems);

        var syntaxTree = MarkdownReader.BuildSyntaxTree(result.Document);
        var definitionTerm = syntaxTree.Children[0].Children[0].Children[0];

        Assert.Equal(MarkdownSyntaxKind.DefinitionTerm, definitionTerm.Kind);
        Assert.Equal("Renamed", definitionTerm.Literal);
    }

    [Fact]
    public void DefinitionList_Parsed_Term_Syntax_Maps_To_Semantic_Term_Node() {
        var result = MarkdownReader.ParseWithSyntaxTree("**Term**: first");
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var term = Assert.Single(group.TermItems);
        var compatibilityTerm = Assert.Single(group.Terms);
        var syntaxTerm = Assert.Single(result.SyntaxTree.Children).Children[0].Children[0];
        var finalSyntaxTerm = Assert.Single(result.FinalSyntaxTree.Children).Children[0].Children[0];

        Assert.Same(term.Inlines, compatibilityTerm);
        Assert.Equal("**Term**", term.Markdown);
        Assert.Equal("Term", term.Text);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 8), term.SourceSpan);
        Assert.Equal(MarkdownSyntaxKind.DefinitionTerm, syntaxTerm.Kind);
        Assert.Same(term, syntaxTerm.AssociatedObject);
        Assert.Same(term, finalSyntaxTerm.AssociatedObject);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void DefinitionList_DefinitionBlockMutation_Rebuilds_Stale_Parsed_SyntaxProjection() {
        var result = MarkdownReader.ParseWithSyntaxTree("Term: first");
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var providedGroup = Assert.Single(definitionList.SyntaxItems);
        var entry = Assert.Single(definitionList.Entries);

        entry.DefinitionBlocks.Clear();
        entry.DefinitionBlocks.Add(new ParagraphBlock(new InlineSequence().Text("second")));

        var ownedGroup = Assert.Single(((IOwnedSyntaxChildrenMarkdownBlock)definitionList).BuildOwnedSyntaxChildren());
        var syntaxTree = MarkdownReader.BuildSyntaxTree(result.Document);
        var rebuiltGroup = Assert.Single(Assert.Single(syntaxTree.Children).Children);
        var definitionMarker = rebuiltGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionMarker);
        var definitionValue = rebuiltGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraph = Assert.Single(definitionValue.Children);

        Assert.NotSame(providedGroup, ownedGroup);
        Assert.NotSame(providedGroup, rebuiltGroup);
        Assert.Equal(":", definitionMarker.Literal);
        Assert.Equal("second", definitionValue.Literal);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Same(entry.DefinitionBlocks[0], paragraph.AssociatedObject);
    }

    private static Markdig.MarkdownPipeline CreateMarkdigDefinitionListPipeline() {
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseDefinitionLists(builder);
        return builder.Build();
    }

    private static Markdig.MarkdownPipeline CreateMarkdigDefinitionListAndPipeTablesPipeline() {
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseDefinitionLists(builder);
        Markdig.MarkdownExtensions.UsePipeTables(builder);
        return builder.Build();
    }

    private static MarkdownReaderOptions CreateMarkdigDefinitionListReaderOptions() {
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.DefinitionLists = true;
        return options;
    }

    private static HtmlOptions CreateMarkdigDefinitionListHtmlOptions() => new() {
        Style = HtmlStyle.Plain,
        CssDelivery = CssDelivery.None,
        BodyClass = null,
        AutoHeadingIdentifiers = false
    };

    private static string NormalizeHtml(string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            return string.Empty;
        }

        var compact = html
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace("> <", "><")
            .Trim();
        var sb = new System.Text.StringBuilder(compact.Length);
        bool lastWasWhitespace = false;
        for (int i = 0; i < compact.Length; i++) {
            char ch = compact[i];
            if (char.IsWhiteSpace(ch)) {
                lastWasWhitespace = true;
                continue;
            }

            if (lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                sb.Append(' ');
            }

            lastWasWhitespace = false;
            sb.Append(ch);
        }

        return sb.ToString();
    }

    private static string NormalizeMarkdown(string markdown) =>
        markdown
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .TrimEnd('\n');

    private static string NormalizePlainText(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        return string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }
}
