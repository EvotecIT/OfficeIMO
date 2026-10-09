using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfTableCompatibilityPlacementTests {
    [Theory]
    [InlineData(WordCompatibilityMode.Word2003, WordTableAlignment.Left, true, 0, 72)]
    [InlineData(WordCompatibilityMode.Word2007, WordTableAlignment.Right, true, 0, 252)]
    [InlineData(WordCompatibilityMode.Word2010, WordTableAlignment.Right, false, 0, 226.8)]
    [InlineData(WordCompatibilityMode.Word2003, WordTableAlignment.Center, true, 720, 162)]
    [InlineData(WordCompatibilityMode.Word2013, WordTableAlignment.Center, false, 720, 149.4)]
    [InlineData(WordCompatibilityMode.Word2013, WordTableAlignment.Right, true, 720, 234)]
    [InlineData(WordCompatibilityMode.Word2013, WordTableAlignment.Left, true, 0, 90)]
    [InlineData(WordCompatibilityMode.Word2013, WordTableAlignment.Left, false, 0, 77.4)]
    public void InlineOriginUsesCompatibilityAndFirstCellMarginWithoutChangingTheGrid(
        WordCompatibilityMode mode, WordTableAlignment alignment, bool firstCellMargin, int indent, double firstX) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        WordTable table = document.AddTable(1, 2);
        table.Alignment = alignment;
        table._tableProperties!.TableIndentation = new TableIndentation { Width = indent, Type = TableWidthUnitValues.Dxa };
        table._tableProperties.TableLayout = new DocumentFormat.OpenXml.Wordprocessing.TableLayout { Type = TableLayoutValues.Fixed };
        table._tableProperties.TableWidth = new TableWidth { Width = "6480", Type = TableWidthUnitValues.Dxa };
        table._tableProperties.TableBorders = new TableBorders(new TopBorder { Val = BorderValues.Nil },
            new LeftBorder { Val = BorderValues.Nil }, new BottomBorder { Val = BorderValues.Nil },
            new RightBorder { Val = BorderValues.Nil }, new InsideHorizontalBorder { Val = BorderValues.Nil },
            new InsideVerticalBorder { Val = BorderValues.Nil });
        foreach (WordTableCell cell in table.Rows[0].Cells) {
            cell.Width = 3240;
            cell.WidthType = WordTableWidthUnit.Dxa;
            cell.MarginLeftWidth = 108;
            cell.MarginRightWidth = 108;
            cell.Paragraphs[0].FontFamily = "Courier New";
            cell.Paragraphs[0].FontSize = 12;
        }
        table.Rows[0].Cells[firstCellMargin ? 0 : 1].MarginLeftWidth = 360;
        table.Rows[0].Cells[0].Paragraphs[0].Text = "FIRST";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "OTHER";
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, FontFamily = "Courier", PageSize = new OfficeIMO.Pdf.PageSize(612, 792),
            Margins = OfficeIMO.Pdf.PageMargins.Uniform(72)
        }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var first = Assert.Single(words, word => word.Text == "FIRST");
        var other = Assert.Single(words, word => word.Text == "OTHER");
        Assert.InRange(Math.Abs(first.Letters[0].StartBaseLine.X - firstX), 0, .03);
        double marginDifference = firstCellMargin ? 5.4 - 18 : 18 - 5.4;
        Assert.InRange(Math.Abs(other.Letters[0].StartBaseLine.X - first.Letters[0].StartBaseLine.X - 162 - marginDifference), 0, .03);
    }
}
