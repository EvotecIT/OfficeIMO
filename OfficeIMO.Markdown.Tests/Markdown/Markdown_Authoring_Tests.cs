using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Authoring_Tests {
    [Fact]
    public void AddMethods_ReturnAttachedBlocksThatLaterEditsRender() {
        MarkdownDocument document = MarkdownDocument.Create();

        HeadingBlock heading = document.AddHeading("Status", 2);
        ParagraphBlock paragraph = document.AddParagraph("Ready");
        OrderedListBlock list = document.AddOrderedList(new[] { "One", "Two" }, start: 3);
        ImageBlock image = document.AddImage("chart.png", altText: "Chart");
        TableBlock table = document.AddTable(t => t.Headers("A", "B").Row("1", "2"));

        list.Start = 5;
        image.Caption = "Figure";

        Assert.Equal(new IMarkdownBlock[] { heading, paragraph, list, image, table }, document.Blocks);
        string markdown = document.ToMarkdown();
        Assert.Contains("## Status", markdown);
        Assert.Contains("5. One", markdown);
        Assert.Contains("![Chart](chart.png)", markdown);
        Assert.Contains("| A | B |", markdown);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(7)]
    public void AddHeading_RejectsLevelsOutsideMarkdownRange(int level) {
        MarkdownDocument document = MarkdownDocument.Create();

        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddHeading("Title", level));
        Assert.Empty(document.Blocks);
    }
}
