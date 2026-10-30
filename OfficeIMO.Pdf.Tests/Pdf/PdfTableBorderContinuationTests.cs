using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableBorderContinuationTests {
    [Fact]
    public void PagedCollapsedBorderRowsKeepEveryRowAfterContinuationPaddingChanges() {
        string[][] rows = Enumerable.Range(0, 24).Select(index => new[] { "Row-" + index.ToString("D2") }).ToArray();
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.AllowRowBreakAcrossPages = false;
        style.CellVerticalPaddingFromBorderInterior = true;
        style.CellPaddingTop = 10D;
        style.CellPaddingBottom = 10D;
        style.CellBorders = Enumerable.Range(0, rows.Length).ToDictionary(
            index => (index, 0),
            _ => new PdfCellBorder {
                Color = PdfColor.Black,
                TopBorder = new PdfCellBorderSide { Color = PdfColor.Black, Width = 1D },
                BottomBorder = new PdfCellBorderSide { Color = PdfColor.Black, Width = 8D }
            });

        byte[] bytes = PdfDocument.Create(new PdfOptions {
            PageWidth = 220D, PageHeight = 140D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
        }).Table(rows, style: style).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);
        string text = string.Join(" ", Enumerable.Range(1, pdf.NumberOfPages).Select(index => pdf.GetPage(index).Text));
        foreach (int index in Enumerable.Range(0, rows.Length)) Assert.Contains("Row-" + index.ToString("D2"), text);
    }
}
