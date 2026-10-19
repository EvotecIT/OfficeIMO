using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_FitsRetainedBannerBeforeChoosingRowBreaks() {
        string path = Path.Combine(_directoryWithFiles, "FitDrawingBounds.xlsx");
        using var workbook = ExcelDocument.Create(path, "Banner");
        var sheet = workbook.Sheets[0];
        sheet.CellValue(1, 1, "Banner");
        sheet.SetColumnWidth(1, 12);
        sheet.SetColumnWidth(2, 12);
        for (int row = 1; row <= 10; row++) {
            sheet.CellValue(row, 2, "Row" + row);
            sheet.SetRowHeight(row, 30);
        }
        sheet.AddImage(1, 1, CreateMinimalRgbPng(), "image/png", widthPixels: 800, heightPixels: 16);
        sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0);
        workbook.Save();
        using var pdf = PdfPigDocument.Open(workbook.ToPdfBytes(new ExcelToPdfOptions {
            PageSize = new PdfCore.PageSize(400, 300), Margins = PdfCore.PageMargins.Uniform(20)
        }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("Row10", pdf.GetPage(1).Text);
        Assert.InRange(FindWordStartY(pdf.GetPage(1), "Row1") - FindWordStartY(pdf.GetPage(1), "Row10"), 161.9, 162.1);
    }
}
