using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, WordTableAlignment.Left, 0, 225D)]
    [InlineData(0, WordTableAlignment.Center, 0, 225D)]
    [InlineData(0, WordTableAlignment.Right, 0, 225D)]
    [InlineData(0, WordTableAlignment.Left, 12, 225D)]
    [InlineData(1, WordTableAlignment.Left, 0, 226.25D)]
    [InlineData(1, WordTableAlignment.Center, 0, 226.25D)]
    [InlineData(1, WordTableAlignment.Right, 0, 226.25D)]
    [InlineData(1, WordTableAlignment.Left, 12, 226.25D)]
    [InlineData(2, WordTableAlignment.Left, 0, 225D)]
    [InlineData(2, WordTableAlignment.Center, 0, 225D)]
    [InlineData(2, WordTableAlignment.Right, 0, 225D)]
    [InlineData(3, WordTableAlignment.Left, 0, 225D)]
    [InlineData(4, WordTableAlignment.Left, 12, 181D)]
    public void SaveAsPdf_FixedTablePreservesAuthoredWidthAndAlignmentBeyondTheTextFrame(
        int widthCase, WordTableAlignment alignment, int indentPoints, double expectedWidth) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.Alignment = alignment;
        table.WidthType = widthCase == 1 ? WordTableWidthUnit.Pct :
            widthCase is 2 or 3 ? WordTableWidthUnit.Auto : WordTableWidthUnit.Dxa;
        table.Width = widthCase == 1 ? 6250 : widthCase is 2 or 3 ? 0 : widthCase == 4 ? 3620 : 4500;
        table.GridColumnWidth = new List<int> { 2250, 2250 };
        table.StyleDetails!.TableIndentationWidth = (short)(indentPoints * 20);
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.WidthType = widthCase == 1 ? WordTableWidthUnit.Pct :
                widthCase == 3 ? WordTableWidthUnit.Auto : WordTableWidthUnit.Dxa;
            cell.Width = widthCase == 1 ? 2500 : widthCase == 3 ? 0 : 2250;
            cell.MarginLeftWidth = 0;
            cell.MarginRightWidth = 0;
            cell.Paragraphs[0].Text = column == 0 ? "L" : "R";
        }
        string originalXml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(325),OfficeIMO.OfficeLength.Points(400)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(72))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters;
        var left = Assert.Single(letters, letter => letter.Value == "L");
        var right = Assert.Single(letters, letter => letter.Value == "R");
        double expectedX = 72 + indentPoints + (alignment == WordTableAlignment.Center
            ? (181 - expectedWidth) / 2 : alignment == WordTableAlignment.Right ? 181 - expectedWidth : 0);
        Assert.Equal(expectedX, left.StartBaseLine.X, 3);
        Assert.Equal(expectedWidth / 2, right.StartBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Equal(originalXml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(WordTableAlignment.Left, 85.5D, 225D)]
    [InlineData(WordTableAlignment.Center, 63D, 225D)]
    [InlineData(WordTableAlignment.Right, 40.5D, 225D)]
    [InlineData(WordTableAlignment.Left, 85.5D, 450D)]
    [InlineData(WordTableAlignment.Center, -49.5D, 450D)]
    [InlineData(WordTableAlignment.Right, -184.5D, 450D)]
    public void SaveAsPdf_FixedTableWidthIncludesTheOuterCellSpacing(WordTableAlignment alignment, double expectedX, double width) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.Alignment = alignment;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = (int)(width * 20D);
        table.GridColumnWidth = new List<int> { 2250, 2250 };
        table.StyleDetails!.CellSpacing = 120;
        table._tableProperties.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.LeftBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.BottomBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.RightBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 8U });
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.WidthType = WordTableWidthUnit.Dxa;
            cell.Width = 2250;
            cell.MarginLeftWidth = cell.MarginRightWidth = 0;
            cell.Paragraphs[0].Text = column == 0 ? "L" : "R";
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(325),OfficeIMO.OfficeLength.Points(400)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(72))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = pdf.GetPage(1);
        var left = Assert.Single(page.Letters, letter => letter.Value == "L");
        var right = Assert.Single(page.Letters, letter => letter.Value == "R");
        Assert.Equal(expectedX, left.StartBaseLine.X, 3);
        Assert.Equal(width / 2D - 6D, right.StartBaseLine.X - left.StartBaseLine.X, 3);
        var borders = page.Paths.Where(path => path.IsStroked)
            .Select(path => path.GetBoundingRectangle()).Where(box => box.HasValue).Select(box => box!.Value).ToArray();
        Assert.Equal(width, borders.Max(box => box.Right) - borders.Min(box => box.Left), 3);
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(4, 0)]
    [InlineData(8, 0)]
    [InlineData(16, 0)]
    [InlineData(0, 108)]
    [InlineData(4, 108)]
    [InlineData(8, 108)]
    [InlineData(16, 108)]
    public void SaveAsPdf_FixedTableSpacingMinimumRetainsTextMarginsAndBorderPaint(int borderSize, int marginTwips) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 600;
        table.GridColumnWidth = new List<int> { 2250, 2250 };
        table.StyleDetails!.CellSpacing = 120;
        W.BorderValues borderValue = borderSize > 0 ? W.BorderValues.Single : W.BorderValues.Nil;
        uint size = (uint)borderSize;
        table._tableProperties.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = borderValue, Size = size },
            new W.LeftBorder { Val = borderValue, Size = size },
            new W.BottomBorder { Val = borderValue, Size = size },
            new W.RightBorder { Val = borderValue, Size = size },
            new W.InsideHorizontalBorder { Val = borderValue, Size = size },
            new W.InsideVerticalBorder { Val = borderValue, Size = size });
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.WidthType = WordTableWidthUnit.Dxa;
            cell.Width = 2250;
            cell.MarginLeftWidth = cell.MarginRightWidth = (short)marginTwips;
            cell.Paragraphs[0].Text = column == 0 ? "L" : "R";
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(325),OfficeIMO.OfficeLength.Points(400)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(72))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = pdf.GetPage(1);
        var left = Assert.Single(page.Letters, letter => letter.Value == "L");
        var right = Assert.Single(page.Letters, letter => letter.Value == "R");
        double borderPoints = borderSize / 8D, marginPoints = marginTwips / 20D;
        Assert.Equal(84D + marginPoints + 1.5D * borderPoints, left.StartBaseLine.X, 3);
        Assert.Equal(13D + 2D * marginPoints + 2D * borderPoints, right.StartBaseLine.X - left.StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(5000, 0, 8, false)]
    [InlineData(6250, 0, 8, false)]
    [InlineData(7500, 0, 8, false)]
    [InlineData(5000, 120, 8, false)]
    [InlineData(6250, 120, 8, false)]
    [InlineData(7500, 120, 8, false)]
    [InlineData(6250, 0, 4, false)]
    [InlineData(6250, 120, 4, false)]
    [InlineData(6250, 0, 16, false)]
    [InlineData(6250, 120, 16, false)]
    [InlineData(6250, 0, 0, true)]
    [InlineData(6250, 120, 0, true)]
    [InlineData(6250, 0, 16, true)]
    [InlineData(6250, 120, 16, true)]
    public void SaveAsPdf_FixedPercentageTableReservesEffectiveOuterBorderPaint(
        int percent, int spacingTwips, int borderSize, bool cellOverride) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Pct;
        table.Width = percent;
        table.GridColumnWidth = new List<int> { 2250, 2250 };
        table.StyleDetails!.CellSpacing = spacingTwips == 0 ? null : (short)spacingTwips;
        uint size = (uint)(cellOverride ? 8 : borderSize);
        W.BorderValues value = size == 0 ? W.BorderValues.Nil : W.BorderValues.Single;
        table._tableProperties.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = value, Size = size },
            new W.LeftBorder { Val = value, Size = size },
            new W.BottomBorder { Val = value, Size = size },
            new W.RightBorder { Val = value, Size = size },
            new W.InsideVerticalBorder { Val = value, Size = size });
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.WidthType = WordTableWidthUnit.Dxa;
            cell.Width = 2250;
            cell.MarginLeftWidth = cell.MarginRightWidth = 0;
            cell.Paragraphs[0].Text = column == 0 ? "L" : "R";
            if (cellOverride) {
                W.BorderValues overrideValue = borderSize == 0 ? W.BorderValues.Nil : W.BorderValues.Single;
                cell._tableCellProperties!.TableCellBorders = column == 0
                    ? new W.TableCellBorders(new W.LeftBorder { Val = overrideValue, Size = (uint)borderSize })
                    : new W.TableCellBorders(new W.RightBorder { Val = overrideValue, Size = (uint)borderSize });
            }
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(325),OfficeIMO.OfficeLength.Points(400)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(72))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters;
        var left = Assert.Single(letters, letter => letter.Value == "L");
        var right = Assert.Single(letters, letter => letter.Value == "R");
        double expectedWidth = (181D - borderSize / 8D) * percent / 5000D;
        double leadingInset = borderSize / 8D * (spacingTwips > 0 ? 1.5D : 1D);
        Assert.InRange(Math.Abs(left.StartBaseLine.X - (72D + spacingTwips / 10D + leadingInset)), 0D, 0.001D);
        double unequalInsets = cellOverride ? (1D - borderSize / 8D) * (spacingTwips > 0 ? 1D : 0.5D) : 0D;
        Assert.InRange(Math.Abs(right.StartBaseLine.X - left.StartBaseLine.X -
            (expectedWidth / 2D - spacingTwips / 20D + unequalInsets)), 0D, 0.001D);
    }
}
