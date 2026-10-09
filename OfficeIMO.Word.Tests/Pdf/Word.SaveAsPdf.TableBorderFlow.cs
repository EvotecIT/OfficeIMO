using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(4U, 0.5D)]
    [InlineData(24U, 3D)]
    public void SaveAsPdf_CollapsedTableBorderPaintFitsBeforeFirstTextAndBeforeFollowingParagraph(uint size, double width) {
        var unbordered = MeasureCollapsedTableFlow(WordBorderStyle.Nil, size);
        var bordered = MeasureCollapsedTableFlow(WordBorderStyle.Single, size);

        // The first cell has one half stroke outside its grid boundary and
        // one half stroke between the boundary and its authored inner margin.
        Assert.Equal(width, unbordered.First - bordered.First, 3);
        // The following paragraph also clears the bottom outward half stroke.
        Assert.Equal(2D * width, unbordered.After - bordered.After, 3);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SaveAsPdf_TableContinuationClearsItsOwnBorderWithoutPaddingFromThePreviousPage(bool splitRow) {
        double thin = MeasureTableContinuation(4U, splitRow);
        double thick = MeasureTableContinuation(24U, splitRow);
        // An incoming border affects the shared boundary on page one. Its
        // clearance does not follow the neighbour onto the next page.
        Assert.Equal(thin, thick, 3);
    }

    private static double MeasureTableContinuation(uint previousBottomSize, bool splitRow) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 1);
        table.ConditionalFormattingFirstRow = false;
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Single, 4U, OfficeIMO.Drawing.OfficeColor.Black);
        for (int row = 0; row < 2; row++) {
            WordTableCell cell = table.Rows[row].Cells[0];
            cell.MarginTopWidth = cell.MarginBottomWidth = 0;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = $"BorderFirst{row}\nBorderNext{row}";
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
            paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
            paragraph.LineSpacing = 240;
        }
        table.Rows[0].Cells[0].Borders.BottomStyle = WordBorderStyle.Single;
        table.Rows[0].Cells[0].Borders.BottomSize = previousBottomSize;
        table.Rows[1].AllowRowToBreakAcrossPages = splitRow;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new OfficeIMO.Pdf.PageSize(300, 90),
            Margins = OfficeIMO.Pdf.PageMargins.Uniform(20)
        }));
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPage(2).GetWords().ToArray();
        return Assert.Single(words, w => w.Text == "BorderNext1").BoundingBox.Bottom;
    }

    private static (double First, double After) MeasureCollapsedTableFlow(WordBorderStyle border, uint size) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.ConditionalFormattingFirstRow = false;
        table.StyleDetails!.SetBordersForAllSides(border, size, OfficeIMO.Drawing.OfficeColor.Black);
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.MarginTopWidth = cell.MarginBottomWidth = 0;
        WordParagraph first = cell.Paragraphs[0];
        first.Text = "BorderFlowFirst";
        WordParagraph after = document.AddParagraph("BorderFlowAfter");
        foreach (WordParagraph paragraph in new[] { first, after }) {
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
            paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
            paragraph.LineSpacing = 240;
        }
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        return (Assert.Single(words, w => w.Text == "BorderFlowFirst").BoundingBox.Bottom,
            Assert.Single(words, w => w.Text == "BorderFlowAfter").BoundingBox.Bottom);
    }
}
