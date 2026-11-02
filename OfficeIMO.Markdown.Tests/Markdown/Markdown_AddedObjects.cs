using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_AddedObjects {
    [Fact]
    public void AddedDefinitionEntry_EditsTheStoredTermAndBody() {
        var list = new DefinitionListBlock();
        var supplied = new DefinitionListEntry(new InlineSequence().Text("Before"),
            new[] { new ParagraphBlock(new InlineSequence().Text("Original")) });
        var entry = list.AddEntry(supplied);
        Assert.Same(entry, Assert.Single(list.Entries));
        entry.Term = new InlineSequence().Text("After");
        entry.DefinitionBlocks.Add(new ParagraphBlock(new InlineSequence().Text("Added body")));
        Assert.Same(entry.Term, Assert.Single(list.Groups).Terms[0]);
        string html = ((IMarkdownBlock)list).RenderHtml();
        Assert.Contains("<dt>After</dt>", html);
        Assert.Contains("Added body", html);
        var parsed = list.AddEntry("Second", "Second body");
        parsed.Term = new InlineSequence().Text("Changed second");
        Assert.Same(parsed, list.Entries[1]);
        Assert.Contains("<dt>Changed second</dt>", ((IMarkdownBlock)list).RenderHtml());
    }

    [Fact]
    public void AddedDefinitionGroup_ReturnsTheStoredGroup() {
        var list = new DefinitionListBlock();
        var group = new DefinitionListGroup(new[] { new InlineSequence().Text("Term") },
            new[] { new DefinitionListDefinition(new[] { new ParagraphBlock(new InlineSequence().Text("Body")) }) });
        Assert.Same(group, list.AddGroup(group));
        Assert.Same(group, Assert.Single(list.Groups));
        Assert.Contains("Body", ((IMarkdownBlock)list).RenderHtml());
    }

    [Fact]
    public void AddedTableRow_RemainsEditableBeforeRendering() {
        var table = new MarkdownTable(new[] { "Name", "Status" });
        var row = table.AddRow("Example", "Draft");
        row[1] = "Reviewed";
        string markdown = ((IMarkdownBlock)table).RenderMarkdown();
        Assert.Contains("Reviewed", markdown);
        Assert.DoesNotContain("Draft", markdown);
        Assert.Contains("Reviewed", ((IMarkdownBlock)table).RenderHtml());
    }
}
