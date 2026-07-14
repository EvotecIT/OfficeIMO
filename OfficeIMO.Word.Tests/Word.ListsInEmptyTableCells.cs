using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
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
        Assert.Equal(NumberFormatValues.Bullet, DocumentTraversal.GetListInfo(first)!.Value.NumberFormat);
        Assert.Equal(1, DocumentTraversal.GetListInfo(nested)!.Value.Level);

        using MemoryStream stream = document.ToStream();
        stream.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(package).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
    }
}
