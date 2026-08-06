using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void TableCell_AddList_PreservesAFormattedEmptyParagraph() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Paragraphs[0].PageBreakBefore = true;

        WordList list = cell.AddList(WordListStyle.Bulleted);
        list.AddItem("First");

        Assert.Equal(2, cell.Paragraphs.Count);
        Assert.True(cell.Paragraphs[0].PageBreakBefore);
        Assert.Equal(string.Empty, cell.Paragraphs[0].Text);
        Assert.Equal("First", cell.Paragraphs[1].Text);
        Assert.True(cell.Paragraphs[1].IsListItem);
    }

    [Fact]
    public void TableCell_HeadingList_KeepsSubsequentItemsInTheCell() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "Existing";

        WordList list = cell.AddList(WordListStyle.Headings111);
        WordParagraph first = list.AddItem("First", 0);
        first.Style = WordParagraphStyles.Heading1;
        WordParagraph second = list.AddItem("Second", 1);
        second.Style = WordParagraphStyles.Heading2;

        Assert.Equal(new[] { "Existing", "First", "Second" }, cell.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        Assert.Same(cell._tableCell, first._paragraph.Parent);
        Assert.Same(cell._tableCell, second._paragraph.Parent);
        Assert.DoesNotContain(document.Sections[0].Paragraphs, paragraph => paragraph.Text == "First" || paragraph.Text == "Second");

        using MemoryStream stream = document.ToStream();
        stream.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(package).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
    }

    [Fact]
    public void TableCell_AddList_WorksAfterItsPlaceholderParagraphIsRemoved() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Paragraphs[0].Remove();
        Assert.Empty(cell.Paragraphs);

        WordList list = cell.AddList(WordListStyle.Bulleted);
        WordParagraph first = list.AddItem("First");
        WordParagraph nested = list.AddItem("Nested", 1);

        Assert.Equal(new[] { "First", "Nested" }, cell.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
        Assert.All(cell.Paragraphs, paragraph => Assert.Same(cell._tableCell, paragraph._paragraph.Parent));
        Assert.Equal(WordNumberFormat.Bullet, WordDocumentTraversal.GetListInfo(first)!.Value.NumberFormat);
        Assert.Equal(1, WordDocumentTraversal.GetListInfo(nested)!.Value.Level);

        using MemoryStream stream = document.ToStream();
        stream.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(package).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
    }
}
