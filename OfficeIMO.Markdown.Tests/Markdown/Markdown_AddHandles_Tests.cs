using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class MarkdownAddHandlesTests {
    [Fact]
    public void Add_Returns_The_Typed_Attached_Block_And_List_Item() {
        using MarkdownDocument document = MarkdownDocument.Create();
        ParagraphBlock paragraph = document.Add(new ParagraphBlock(new InlineSequence().Text("Before")));
        Assert.Same(paragraph, document.Blocks[0]);
        paragraph.Inlines.Text("after");
        MarkdownUnorderedList list = document.Add(new MarkdownUnorderedList());
        ListItem item = ListItem.Text("Item");
        Assert.Same(item, list.Add(item));
        item.Content.Text("edited");
        Assert.Equal("Before after\n\n- Item edited\n", document.ToMarkdown().Replace("\r\n", "\n"));
        Assert.Throws<ArgumentNullException>(() => document.Add<IMarkdownBlock>(null!));
        Assert.Throws<ArgumentNullException>(() => list.Add(null!));
        Assert.Equal(2, document.Blocks.Count);
    }

    [Fact]
    public void AddRange_Validates_The_Entire_Input_Before_Changing_A_Parsed_Document() {
        using MarkdownDocument document = MarkdownDocument.Parse("Original\n");
        IMarkdownBlock original = Assert.Single(document.Blocks);
        var candidate = new ParagraphBlock(new InlineSequence().Text("Candidate"));
        Assert.Throws<ArgumentException>(() => document.AddRange(new IMarkdownBlock[] { candidate, null! }));
        Assert.Throws<InvalidOperationException>(() => document.AddRange(FailingInput(candidate)));
        Assert.Same(original, Assert.Single(document.Blocks));
        Assert.Null(candidate.Parent);
        Assert.Equal("Original\n", document.ToMarkdown().Replace("\r\n", "\n"));
    }

    private static IEnumerable<IMarkdownBlock> FailingInput(IMarkdownBlock first) {
        yield return first;
        throw new InvalidOperationException("The source enumeration failed.");
    }
}
