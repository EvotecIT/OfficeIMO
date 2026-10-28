using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class MarkdownDocBatchTests {
    [Fact]
    public void AddRange_Appends_In_Order_And_Binds_The_Object_Tree() {
        var first = new ParagraphBlock(new InlineSequence().Text("First"));
        var second = new ParagraphBlock(new InlineSequence().Text("Second"));
        MarkdownDocument document = MarkdownDocument.Create();

        IReadOnlyList<IMarkdownBlock> returned = document.AddRange(new IMarkdownBlock[] { first, second });

        Assert.Equal(new IMarkdownBlock[] { first, second }, returned);
        Assert.Same(first, returned[0]);
        Assert.Same(second, returned[1]);
        Assert.Equal(new IMarkdownBlock[] { first, second }, document.Blocks);
        Assert.Same(document, first.Parent);
        Assert.Same(document, second.Parent);
        Assert.Same(second, first.NextSibling);
        Assert.Same(first, second.PreviousSibling);
        Assert.Equal("First\n\nSecond\n", document.ToMarkdown().Replace("\r\n", "\n"));
    }

    [Fact]
    public void AddRange_Rejects_Null_Input_And_Null_Entries() {
        MarkdownDocument document = MarkdownDocument.Create();

        Assert.Throws<ArgumentNullException>(() => document.AddRange(null!));
        Assert.Throws<ArgumentException>(() => document.AddRange(new IMarkdownBlock[] { null! }));
    }
}
