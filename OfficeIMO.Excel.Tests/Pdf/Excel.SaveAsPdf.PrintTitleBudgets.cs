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
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1600, 1, "End");
        document.SetPrintArea(sheet, "A1:A1600");
        document.SetPrintTitles(sheet, firstRow: 3, lastRow: 1599, firstCol: null, lastCol: null);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, PageSize = new PdfCore.PdfPageSize(100, 100),
            Margins = PdfCore.PdfPageMargins.Uniform(10)
        }).ToBytes());
        Assert.Contains("page projection", error.Message);
    }

    [Fact]
    public void SaveAsPdf_DenseMiddleTitleColumnsRejectsExcessivePageProjection() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1600, "End");
        document.SetPrintArea(sheet, "A1:BIN1");
        document.SetPrintTitles(sheet, firstRow: null, lastRow: null, firstCol: 3, lastCol: 1599);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, PageSize = new PdfCore.PdfPageSize(100, 100),
            Margins = PdfCore.PdfPageMargins.Uniform(10)
        }).ToBytes());
        Assert.Contains("page projection", error.Message);
    }

    [Fact]
    public void SaveAsPdf_FlowTableBoundsCombinedManualBreakTitleProjection() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(50, 50, "End");
        document.SetPrintArea(sheet, "A1:AX50");
        document.SetPrintTitles(sheet, firstRow: 3, lastRow: 40, firstCol: 3, lastCol: 40);
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
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(2, 2, "End");
        document.SetPrintArea(sheet, "A1:B2");
        document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1);
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
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, "Merged");
        sheet.Cell(11, 1, "Body");
        sheet.MergeRange("A1:A10");
        document.SetPrintArea(sheet, "A11:A11");
        document.SetPrintTitles(sheet, firstRow: 1, lastRow: 10, firstCol: null, lastCol: null);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, MaxRowsPerSheet = 5,
            PageSize = new PdfCore.PdfPageSize(300, 300), Margins = PdfCore.PdfPageMargins.Uniform(20)
        }).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        Assert.True(pdf.NumberOfPages > 0);
    }
}
