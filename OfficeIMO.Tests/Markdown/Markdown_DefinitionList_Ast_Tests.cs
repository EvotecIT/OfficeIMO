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
    public void DefinitionList_MarkdigLazyContinuation_Stays_In_Definition_Paragraph_Source() {
        const string markdown = """
Term
:   First paragraph
lazy continuation
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(result.Document.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var definition = Assert.Single(group.Definitions);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(definition.Blocks));
        var syntaxGroup = Assert.Single(result.SyntaxTree.Children).Children[0];
        var definitionValue = syntaxGroup.Children.Single(child => child.Kind == MarkdownSyntaxKind.DefinitionValue);
        var paragraphSyntax = Assert.Single(definitionValue.Children);

        Assert.Equal("First paragraph lazy continuation", paragraph.Inlines.RenderMarkdown());
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 17), definitionValue.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 5, 3, 17), paragraphSyntax.SourceSpan);
        Assert.Same(definition, definitionValue.AssociatedObject);
        Assert.Same(paragraph, paragraphSyntax.AssociatedObject);
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
}
