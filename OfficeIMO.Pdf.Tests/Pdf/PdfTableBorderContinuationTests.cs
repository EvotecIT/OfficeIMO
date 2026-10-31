using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableBorderContinuationTests {
    [Fact]
    public void PagedCollapsedBorderRowsKeepEveryRowAfterContinuationPaddingChanges() {
        string[][] rows = Enumerable.Range(0, 24).Select(index => new[] { "Row-" + index.ToString("D2") }).ToArray();
        var style = PdfTableStyles.Minimal();
        style.BorderColor = null;
        style.AllowRowBreakAcrossPages = false;
        style.CellPaddingFromBorderInterior = true;
        style.CellPaddingTop = OfficeLength.Points(10D);
        style.CellPaddingBottom = OfficeLength.Points(10D);
        style.CellBorders = Enumerable.Range(0, rows.Length).ToDictionary(
            index => (index, 0),
            _ => new PdfCellBorder {
                Color = OfficeColor.Black,
                TopBorder = new PdfCellBorderSide { Color = OfficeColor.Black, Width = OfficeLength.Points(1D) },
                BottomBorder = new PdfCellBorderSide { Color = OfficeColor.Black, Width = OfficeLength.Points(8D) }
            });

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = OfficeLength.Points(220D), PageHeight = OfficeLength.Points(140D),
            MarginLeft = OfficeLength.Points(20D), MarginRight = OfficeLength.Points(20D), MarginTop = OfficeLength.Points(20D), MarginBottom = OfficeLength.Points(20D)
        }).Table(rows, style: style).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);
        string text = string.Join(" ", Enumerable.Range(1, pdf.NumberOfPages).Select(index => pdf.GetPage(index).Text));
        foreach (int index in Enumerable.Range(0, rows.Length)) Assert.Contains("Row-" + index.ToString("D2"), text);
    }
}
