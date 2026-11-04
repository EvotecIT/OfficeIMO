using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_WrappedAndShrunkCellsPreserveNeighborFontSizes() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Layout");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(12));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(25));
        sheet.SetColumnWidth(3, ExcelColumnWidth.Characters(15));
        sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(45)));
        sheet.SetRowHeight(2, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(24)));
        sheet.Cell(1, 1).SetValue("Neighbor1");
        sheet.Cell(1, 2).SetValue("The quick brown fox jumps over the lazy dog with enough words to wrap.");
        sheet.Cell(1, 3).SetValue("Amount");
        sheet.Cell(1, 2).SetWrapText();
        sheet.Cell(2, 1).SetValue("Neighbor2");
        sheet.Cell(2, 2).SetValue("Normal");
        sheet.Cell(2, 3).SetValue("Shrink this long cell text");
        sheet.Cell(2, 3).SetShrinkToFit();
        for (int row = 1; row <= 2; row++) {
            for (int column = 1; column <= 3; column++) {
                sheet.Cell(row, column).SetFontSize(global::OfficeIMO.OfficeLength.Points(11));
                sheet.Cell(row, column).SetFontName("Arial");
            }
        }
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, IncludeSheetHeadings = false
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        foreach (string text in new[] { "Neighbor1", "Neighbor2", "Amount", "Normal", "quick", "wrap." }) {
            var word = Assert.Single(words, word => word.Text == text);
            Assert.All(word.Letters, letter => Assert.InRange(letter.PointSize, 10.99, 11.01));
        }
        var quick = Assert.Single(words, word => word.Text == "quick");
        var wrap = Assert.Single(words, word => word.Text == "wrap.");
        Assert.True(quick.BoundingBox.Bottom - wrap.BoundingBox.Bottom > 15);
        var shrunk = Assert.Single(words, word => word.Text == "Shrink");
        Assert.All(shrunk.Letters, letter => Assert.InRange(letter.PointSize, 4, 10));
    }
}
