using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordBorderStyle.Single)]
    [InlineData(WordBorderStyle.Double)]
    public void SaveAsPdf_BorderedTableHeightKeepsThePrecedingHeadingWithItsRow(WordBorderStyle borderStyle) {
        using WordDocument document = WordDocument.Create();
        ConfigurePartialColumnParagraph(document.AddParagraph("BeforeOne\nBeforeTwo\nBeforeThree"));
        var heading = document.AddParagraph("BorderHeading");
        ConfigurePartialColumnParagraph(heading);
        heading.KeepWithNextOverride = true;
        WordTable table = document.AddTable(1, 1);
        ConfigureMeasuredBorderTable(table, borderStyle);
        var cell = table.Rows[0].Cells[0];
        ConfigurePartialColumnParagraph(cell.Paragraphs[0]);
        cell.Paragraphs[0].Text = "BorderRow";

        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false,
            PageSize = new OfficeIMO.Pdf.PageSize(300, 140),
            Margins = OfficeIMO.Pdf.PageMargins.Uniform(20)
        }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("BeforeThree", pdf.GetPage(1).Text);
        Assert.DoesNotContain("BorderHeading", pdf.GetPage(1).Text);
        Assert.Contains("BorderHeading", pdf.GetPage(2).Text);
        Assert.Contains("BorderRow", pdf.GetPage(2).Text);
    }

    private static void ConfigureMeasuredBorderTable(WordTable table, WordBorderStyle style) {
        table.ConditionalFormattingFirstRow = false;
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.Width = 4000; table.WidthType = WordTableWidthUnit.Dxa;
        table.StyleDetails!.SetBordersForAllSides(style, 12U, OfficeIMO.Drawing.OfficeColor.Black);
        foreach (var row in table.Rows) {
            var cell = row.Cells[0];
            cell.Width = 4000; cell.WidthType = WordTableWidthUnit.Dxa;
            cell.MarginTopWidth = 0; cell.MarginBottomWidth = 0;
        }
    }

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
