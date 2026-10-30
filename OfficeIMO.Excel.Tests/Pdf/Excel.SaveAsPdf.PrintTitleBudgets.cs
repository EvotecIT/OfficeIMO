using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
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
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToPdfBytes(new ExcelToPdfOptions {
            HeaderRowCount = 0, PageSize = new PdfCore.PageSize(100, 100),
            Margins = PdfCore.PageMargins.Uniform(10)
        }));
        Assert.Contains("page projection", error.Message);
    }

    [Fact]
    public void SaveAsPdf_DenseMiddleTitleColumnsRejectsExcessivePageProjection() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1600, "End");
        document.SetPrintArea(sheet, "A1:BIN1");
        document.SetPrintTitles(sheet, firstRow: null, lastRow: null, firstCol: 3, lastCol: 1599);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => document.ToPdfBytes(new ExcelToPdfOptions {
            HeaderRowCount = 0, PageSize = new PdfCore.PageSize(100, 100),
            Margins = PdfCore.PageMargins.Uniform(10)
        }));
        Assert.Contains("page projection", error.Message);
    }
}
