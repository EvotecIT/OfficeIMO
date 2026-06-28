using OfficeIMO.Markdown;
using MarkdigMarkdown = Markdig.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_DefinitionList_Ast_Tests {
    [Theory]
    [InlineData("Term\n:   Definition\n")]
    [InlineData("Term\n:\tDefinition\n")]
    public void DefinitionList_MarkdigMarkerSyntax_Matches_MarkdigHtml(string markdown) {
        var office = MarkdownReader.Parse(markdown).ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });
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
                MarkdownSyntaxKind.DefinitionValue,
                MarkdownSyntaxKind.DefinitionValue
            },
            syntaxGroup.Children.Select(child => child.Kind).ToArray());
        Assert.Same(group.TermItems[0], syntaxGroup.Children[0].AssociatedObject);
        Assert.Same(group.TermItems[1], syntaxGroup.Children[1].AssociatedObject);
        Assert.Same(group.Definitions[0], syntaxGroup.Children[2].AssociatedObject);
        Assert.Same(group.Definitions[1], syntaxGroup.Children[3].AssociatedObject);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [InlineData("Term\n: Definition\n")]
    [InlineData("Term\n:  Definition\n")]
    public void DefinitionList_MarkdigMarkerSyntax_Requires_MarkdigMarkerSpacing(string markdown) {
        var doc = MarkdownReader.Parse(markdown);
        var html = doc.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });
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
        var definitionValue = rebuiltGroup.Children[1];
        var paragraph = Assert.Single(definitionValue.Children);

        Assert.NotSame(providedGroup, ownedGroup);
        Assert.NotSame(providedGroup, rebuiltGroup);
        Assert.Equal("second", definitionValue.Literal);
        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Same(entry.DefinitionBlocks[0], paragraph.AssociatedObject);
    }

    private static Markdig.MarkdownPipeline CreateMarkdigDefinitionListPipeline() {
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseDefinitionLists(builder);
        return builder.Build();
    }

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
