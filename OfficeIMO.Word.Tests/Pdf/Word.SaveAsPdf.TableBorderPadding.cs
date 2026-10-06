using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordBorderStyle.Single, 12U, 0, 1.5D)]
    [InlineData(WordBorderStyle.Single, 12U, 60, 1.5D)]
    [InlineData(WordBorderStyle.Double, 12U, 0, 4.5D)]
    [InlineData(WordBorderStyle.Double, 12U, 60, 4.5D)]
    [InlineData(WordBorderStyle.Double, 24U, 60, 9D)]
    public void SaveAsPdf_TableBorders_AddPaintClearanceToAuthoredVerticalMargins(
        WordBorderStyle style, uint size, short marginTwips, double expectedAdditionalHeight) {
        double unbordered = GetBorderedTableRowDistance(WordBorderStyle.Nil, size, marginTwips);
        double bordered = GetBorderedTableRowDistance(style, size, marginTwips);

        // Word measures vertical cell margins from the border interior. Paint
        // must add to those margins even when the authored margin is already
        // larger than the clearance, rather than being absorbed by it.
        Assert.Equal(expectedAdditionalHeight, bordered - unbordered, 3);
    }

    private static double GetBorderedTableRowDistance(WordBorderStyle style, uint size, short marginTwips) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 1);
        table.ConditionalFormattingFirstRow = false;
        table.StyleDetails!.SetBordersForAllSides(style, size, OfficeIMO.Drawing.OfficeColor.Black);
        for (int row = 0; row < 2; row++) {
            var cell = table.Rows[row].Cells[0];
            cell.MarginTopWidth = marginTwips;
            cell.MarginBottomWidth = marginTwips;
            var paragraph = cell.Paragraphs[0];
            paragraph.Text = $"RowMarker{row}";
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = 0;
            paragraph.LineSpacingAfterPoints = 0;
            paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
            paragraph.LineSpacing = 240;
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var first = Assert.Single(words, word => word.Text == "RowMarker0");
        var second = Assert.Single(words, word => word.Text == "RowMarker1");
        return first.BoundingBox.Bottom - second.BoundingBox.Bottom;
    }
}
