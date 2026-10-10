using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(1200)]
    [InlineData(6000)]
    public void SaveAsPdf_AutomaticCellWidthsUseUnwrappedContentInsteadOfCachedGrid(int gridTwips) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2, WordTableStyle.TableGrid);
        table.WidthType = WordTableWidthUnit.Auto;
        table.Width = 0;
        table.GridColumnWidth = new List<int> { gridTwips, gridTwips };
        table._tableProperties.TableLayout = new W.TableLayout { Type = W.TableLayoutValues.Autofit };
        table._tableProperties.TableCellMarginDefault = new W.TableCellMarginDefault(
            new W.TopMargin { Type = W.TableWidthUnitValues.Dxa, Width = "0" },
            new W.BottomMargin { Type = W.TableWidthUnitValues.Dxa, Width = "0" },
            new W.TableCellLeftMargin { Type = W.TableWidthValues.Dxa, Width = 240 },
            new W.TableCellRightMargin { Type = W.TableWidthValues.Dxa, Width = 240 });
        string[,] text = { { "Alpha Bravo", "Z" }, { "Row two", "Second cell" } };
        for (int row = 0; row < 2; row++) {
            for (int column = 0; column < 2; column++) {
                WordTableCell cell = table.Rows[row].Cells[column];
                cell.WidthType = WordTableWidthUnit.Auto;
                cell.Width = 0;
                WordParagraph paragraph = cell.Paragraphs[0];
                paragraph.Text = text[row, column];
                paragraph.FontFamily = "Courier New";
                paragraph.FontSize = 12;
                paragraph.LineSpacingBeforePoints = 0;
                paragraph.LineSpacingAfterPoints = 0;
            }
        }
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = new PdfOptions { DefaultFont = PdfStandardFont.Courier }
        }));
        var letters = pdf.GetPage(1).Letters;
        var alpha = Assert.Single(letters, x => x.Value == "A");
        var bravo = Assert.Single(letters, x => x.Value == "B");
        var rowStart = Assert.Single(letters, x => x.Value == "R");
        var second = Assert.Single(letters, x => x.Value == "S");
        var lastLetter = letters.Last(x => x.Value == "l");
        Assert.Equal(alpha.StartBaseLine.Y, bravo.StartBaseLine.Y, 3);
        Assert.Equal(rowStart.StartBaseLine.Y, second.StartBaseLine.Y, 3);
        Assert.Equal(second.StartBaseLine.Y, lastLetter.StartBaseLine.Y, 3);
        // Both columns need the same longest line, despite different cached grids.
        Assert.InRange(second.StartBaseLine.X - rowStart.StartBaseLine.X, 100D, 110D);
    }
}
