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
        using var workbook = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Banner" });
        var sheet = workbook.Sheets[0];
        sheet.CellValue(1, 1, "Banner");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(12));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(12));
        for (int row = 1; row <= 10; row++) {
            sheet.CellValue(row, 2, "Row" + row);
            sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
        }
        sheet.AddImage(1, 1, CreateMinimalRgbPng(), new OfficeIMO.Excel.ExcelImageOptions {
            ContentType = "image/png",
            Width = OfficeIMO.OfficeLength.Pixels(800),
            Height = OfficeIMO.OfficeLength.Pixels(16)
        });
        sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0);
        workbook.Save();
        using var pdf = PdfPigDocument.Open(workbook.ToPdf(new ExcelPdfExportOptions {
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(400),OfficeIMO.OfficeLength.Points(300)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("Row10", pdf.GetPage(1).Text);
        Assert.InRange(FindWordStartY(pdf.GetPage(1), "Row1") - FindWordStartY(pdf.GetPage(1), "Row10"), 161.9, 162.1);
    }
}
