using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(1200, 1200, 0, false)]
    [InlineData(6000, 6000, 0, false)]
    [InlineData(1200, 6000, 0, false)]
    [InlineData(6000, 1200, 0, false)]
    [InlineData(1200, 1200, 120, false)]
    [InlineData(1200, 6000, 120, false)]
    [InlineData(1200, 1200, 0, true)]
    public void SaveAsPdf_AutomaticCellWidthsUseUnwrappedContentInsteadOfCachedGrid(int firstGridTwips, int secondGridTwips, int spacingTwips, bool longWord) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 2, WordTableStyle.TableGrid);
        table.WidthType = WordTableWidthUnit.Auto;
        table.Width = 0;
        table.GridColumnWidth = new List<int> { firstGridTwips, secondGridTwips };
        table._tableProperties.TableLayout = new W.TableLayout { Type = W.TableLayoutValues.Autofit };
        if (spacingTwips > 0) {
            table._tableProperties.TableCellSpacing = new W.TableCellSpacing {
                Type = W.TableWidthUnitValues.Dxa, Width = spacingTwips.ToString(System.Globalization.CultureInfo.InvariantCulture)
            };
        }
        table._tableProperties.TableCellMarginDefault = new W.TableCellMarginDefault(
            new W.TopMargin { Type = W.TableWidthUnitValues.Dxa, Width = "0" },
            new W.BottomMargin { Type = W.TableWidthUnitValues.Dxa, Width = "0" },
            new W.TableCellLeftMargin { Type = W.TableWidthValues.Dxa, Width = 240 },
            new W.TableCellRightMargin { Type = W.TableWidthValues.Dxa, Width = 240 });
        string[,] text = { { longWord ? "AlphaBetaGammaDeltaEpsilon" : "Alpha Bravo", "Z" }, { "Row two", "Second cell" } };
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
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = new PdfCreateOptions { DefaultFont = PdfStandardFont.Courier }
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters;
        var alpha = Assert.Single(letters, x => x.Value == "A");
        var bravo = Assert.Single(letters, x => x.Value == "B");
        var rowStart = Assert.Single(letters, x => x.Value == "R");
        var second = Assert.Single(letters, x => x.Value == "S");
        var lastLetter = letters.Last(x => x.Value == "l");
        Assert.Equal(alpha.StartBaseLine.Y, bravo.StartBaseLine.Y, 3);
        Assert.Equal(rowStart.StartBaseLine.Y, second.StartBaseLine.Y, 3);
        Assert.Equal(second.StartBaseLine.Y, lastLetter.StartBaseLine.Y, 3);
        // Content and the independent border frame determine the column origin.
        // Cached grids must not take width from either unwrapped line.
        double expectedColumnWidth = longWord ? 211.2D : 103.2D;
        double spacingFrameWidth = spacingTwips > 0 ? 13D : 0D;
        Assert.InRange(second.StartBaseLine.X - rowStart.StartBaseLine.X,
            expectedColumnWidth + spacingFrameWidth - 1D, expectedColumnWidth + spacingFrameWidth + 1D);
    }
}
