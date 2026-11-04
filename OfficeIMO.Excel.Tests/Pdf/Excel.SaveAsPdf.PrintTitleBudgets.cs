using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_DenseMiddleTitleRowsRejectsExcessivePageProjection() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1600, 1).SetValue("End");
        sheet.PageSetup.PrintArea = "A1:A1600";
        sheet.PageSetup.SetPrintTitles(firstRow: 3, lastRow: 1599, firstColumn: null, lastColumn: null);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, PageSize = new PdfCore.PdfPageSize(OfficeLength.Points(100), OfficeLength.Points(100)),
            Margins = PdfCore.PdfPageMargins.Uniform(OfficeLength.Points(10))
        }).ToBytes());
        Assert.Contains("page projection", error.Message);
    }

    [Fact]
    public void SaveAsPdf_DenseMiddleTitleColumnsRejectsExcessivePageProjection() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1, 1600).SetValue("End");
        sheet.PageSetup.PrintArea = "A1:BIN1";
        sheet.PageSetup.SetPrintTitles(firstRow: null, lastRow: null, firstColumn: 3, lastColumn: 1599);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, PageSize = new PdfCore.PdfPageSize(OfficeLength.Points(100), OfficeLength.Points(100)),
            Margins = PdfCore.PdfPageMargins.Uniform(OfficeLength.Points(10))
        }).ToBytes());
        Assert.Contains("page projection", error.Message);
    }

    [Fact]
    public void SaveAsPdf_FlowTableBoundsCombinedManualBreakTitleProjection() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(50, 50).SetValue("End");
        sheet.PageSetup.PrintArea = "A1:AX50";
        sheet.PageSetup.SetPrintTitles(firstRow: 3, lastRow: 40, firstColumn: 3, lastColumn: 40);
        for (int index = 1; index < 50; index++) {
            sheet.AddManualRowPageBreak(index);
            sheet.AddManualColumnPageBreak(index);
        }
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        }).ToBytes());
        Assert.Contains("page projection", error.Message);
    }

    [Fact]
    public void SaveAsPdf_ProjectionUsesConfiguredWorksheetCellBudget() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(2, 2).SetValue("End");
        sheet.PageSetup.PrintArea = "A1:B2";
        sheet.PageSetup.SetPrintTitles(firstRow: 1, lastRow: 1, firstColumn: 1, lastColumn: 1);
        sheet.AddManualRowPageBreak(1);
        sheet.AddManualColumnPageBreak(1);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable,
            MaxMaterializedWorksheetCells = 5
        }).ToBytes());
        Assert.Contains("page projection", error.Message);
    }

    [Fact]
    public void SaveAsPdf_MergedPrependedTitleRetainsFullGeometryWhenBodyRowsAreCapped() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("Merged");
        sheet.Cell(11, 1).SetValue("Body");
        sheet.Range("A1:A10").Merge();
        sheet.PageSetup.PrintArea = "A11:A11";
        sheet.PageSetup.SetPrintTitles(firstRow: 1, lastRow: 10, firstColumn: null, lastColumn: null);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, MaxRowsPerSheet = 5,
            PageSize = new PdfCore.PdfPageSize(OfficeLength.Points(300), OfficeLength.Points(300)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeLength.Points(20))
        }).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        Assert.True(pdf.NumberOfPages > 0);
    }
}
