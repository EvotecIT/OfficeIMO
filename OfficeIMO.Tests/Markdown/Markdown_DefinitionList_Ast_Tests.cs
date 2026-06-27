using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_DefinitionList_Ast_Tests {
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
        Assert.All(definitionList.Entries, entry => Assert.Equal("Renamed", entry.TermMarkdown));
        Assert.Equal("Renamed: first\nRenamed: second", ((IMarkdownBlock)definitionList).RenderMarkdown());

        var html = ((IMarkdownBlock)definitionList).RenderHtml();
        Assert.Contains("<dt>Renamed</dt>", html);
        Assert.DoesNotContain("Original", html);
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
}
