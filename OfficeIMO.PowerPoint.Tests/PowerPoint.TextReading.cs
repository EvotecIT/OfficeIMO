using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointTextReadingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SavedTextRetainsInlineBreaksFieldTextAndParagraphBoundaries(bool tableCell) {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        PowerPointParagraph paragraph;
        if (tableCell) {
            var cell = slide.AddTable(1, 1).GetCell(0, 0);
            cell.Text = "First";
            paragraph = cell.Paragraphs[0];
            cell.AddParagraph("Last");
        } else {
            var box = slide.AddTextBox("First");
            paragraph = box.Paragraphs[0];
            box.AddParagraph("Last");
        }
        paragraph.AddLineBreak().AddField("2", "slidenum").AddText(" tail");
        Assert.Equal("First" + Environment.NewLine + "2 tail", paragraph.Text);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        string actual = tableCell ? Assert.Single(saved.Slides[0].Tables).GetCell(0, 0).Text : Assert.Single(saved.Slides[0].TextBoxes).Text;
        Assert.Equal("First" + Environment.NewLine + "2 tail" + Environment.NewLine + "Last", actual);
        Assert.Empty(saved.ValidateDocument());
    }
}
