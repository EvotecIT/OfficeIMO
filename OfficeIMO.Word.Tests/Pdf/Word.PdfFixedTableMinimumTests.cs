using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, 0, 600, 84D, 13D)]
    [InlineData(0, 0, 960, 84D, 13D)]
    [InlineData(8, 0, 600, 85.5D, 15D)]
    [InlineData(8, 108, 600, 90.9D, 25.8D)]
    [InlineData(8, 0, 4500, 85.5D, 15D)]
    public void SaveAsPdf_UnequalFixedColumnsRetainPhysicalMinimums(
        int border, short margin, int tableWidth, double expectedLeft, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateUnequalFixedMinimumControl(document, 1, 20, 120);
        table.Width = tableWidth;
        W.BorderValues value = border == 0 ? W.BorderValues.Nil : W.BorderValues.Single;
        table._tableProperties.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = value, Size = (uint)border },
            new W.LeftBorder { Val = value, Size = (uint)border },
            new W.BottomBorder { Val = value, Size = (uint)border },
            new W.RightBorder { Val = value, Size = (uint)border },
            new W.InsideVerticalBorder { Val = value, Size = (uint)border });
        foreach (WordTableCell cell in table.Rows[0].Cells)
            cell.MarginLeftWidth = cell.MarginRightWidth = margin;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(UnequalFixedMinimumOptions()));
        var letters = pdf.GetPage(1).Letters;
        double left = Assert.Single(letters, letter => letter.Value == "L").StartBaseLine.X;
        double right = Assert.Single(letters, letter => letter.Value == "R").StartBaseLine.X;
        Assert.Equal(expectedLeft, left, 3);
        Assert.Equal(expectedGap, right - left, 3);
    }

    [Theory]
    [InlineData(20, 120, 13D)]
    [InlineData(4480, 120, 200D)]
    [InlineData(20, 0, 1D)]
    [InlineData(4480, 0, 224D)]
    public void SaveAsPdf_MergedWidthAndMarginsPreserveTheUnderlyingUnequalGrid(int first, short spacing, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateUnequalFixedMinimumControl(document, 2, first, spacing);
        table.Rows[0].Cells[0].MergeHorizontally(1);
        WordTableCell merged = table.Rows[0].Cells[0];
        merged.Width = 4500;
        merged.MarginLeftWidth = merged.MarginRightWidth = 400;
        merged.Paragraphs[0].Text = "Span";
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(UnequalFixedMinimumOptions()));
        var letters = pdf.GetPage(1).Letters;
        double left = Assert.Single(letters, letter => letter.Value == "L").StartBaseLine.X;
        double right = Assert.Single(letters, letter => letter.Value == "R").StartBaseLine.X;
        Assert.Equal(72D + spacing / 10D, left, 3);
        Assert.Equal(expectedGap, right - left, 3);
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(0, 4000, 0, 20, 2D / 227D * 401D)]
    [InlineData(120, 4000, 0, 20, 2D / 227D * 389D + 12D)]
    [InlineData(0, 4000, 0, 1125, 57.25D / 227D * 401D)]
    [InlineData(120, 4000, 0, 1125, 57.25D / 227D * 389D + 12D)]
    [InlineData(0, 2400, 400, 20, 21D)]
    [InlineData(0, 3000, 400, 20, 82D / 307D * 301D - 20D)]
    public void SaveAsPdf_MergedMarginsExpandAFixedTableToRetainItsTextFrame(
        short spacing, short mergedMargin, short cellMargin, int first, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateUnequalFixedMinimumControl(document, 2, first, spacing);
        table.Rows[1].Cells[0].MarginLeftWidth = table.Rows[1].Cells[0].MarginRightWidth = cellMargin;
        table.Rows[0].Cells[0].MergeHorizontally(1);
        WordTableCell merged = table.Rows[0].Cells[0];
        merged.Width = 4500;
        merged.MarginLeftWidth = merged.MarginRightWidth = mergedMargin;
        merged.Paragraphs[0].Text = "S";
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(UnequalFixedMinimumOptions()));
        var letters = pdf.GetPage(1).Letters;
        double x = Assert.Single(letters, letter => letter.Value == "S").StartBaseLine.X;
        Assert.Equal(72D + mergedMargin / 20D + spacing / 10D, x, 3);
        double left = Assert.Single(letters, letter => letter.Value == "L").StartBaseLine.X;
        double right = Assert.Single(letters, letter => letter.Value == "R").StartBaseLine.X;
        Assert.Equal(expectedGap, right - left, 3);
        Assert.Equal(1, pdf.NumberOfPages);
    }

    private static WordTable CreateUnequalFixedMinimumControl(WordDocument document, int rows, int first, short spacing) {
        WordTable table = document.AddTable(rows, 2, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = 4500;
        table.GridColumnWidth = new List<int> { first, 4500 - first };
        table.StyleDetails!.CellSpacing = spacing;
        table._tableProperties.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.Nil }, new W.LeftBorder { Val = W.BorderValues.Nil },
            new W.BottomBorder { Val = W.BorderValues.Nil }, new W.RightBorder { Val = W.BorderValues.Nil },
            new W.InsideHorizontalBorder { Val = W.BorderValues.Nil }, new W.InsideVerticalBorder { Val = W.BorderValues.Nil });
        foreach (WordTableRow row in table.Rows)
            for (int column = 0; column < 2; column++) {
                WordTableCell cell = row.Cells[column];
                cell.WidthType = WordTableWidthUnit.Dxa;
                cell.Width = column == 0 ? first : 4500 - first;
                cell.MarginLeftWidth = cell.MarginRightWidth = 0;
                cell.Paragraphs[0].Text = column == 0 ? "L" : "R";
            }
        return table;
    }

    private static WordToPdfOptions UnequalFixedMinimumOptions() => new() {
        IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
        PageSize = new PageSize(600, 400), Margins = PageMargins.Uniform(72)
    };
}
