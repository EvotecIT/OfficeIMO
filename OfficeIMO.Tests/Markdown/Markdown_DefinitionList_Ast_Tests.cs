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
}
