using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelPageOrder.DownThenOver)]
    [InlineData(ExcelPageOrder.OverThenDown)]
    public void SaveAsPdf_TitleOnlyManualPageKeepsItsSourceOrder(ExcelPageOrder pageOrder) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        for (int row = 1; row <= 8; row++) sheet.Cell(row, 1, "Row" + row.ToString("D3"));
        document.SetPrintArea(sheet, "A1:A8");
        document.SetPrintTitles(sheet, firstRow: 3, lastRow: 4, firstCol: null, lastCol: null);
        sheet.SetPageSetup(pageOrder: pageOrder);
        sheet.AddManualRowPageBreak(2); sheet.AddManualRowPageBreak(4);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false,
            PageSize = new PdfCore.PageSize(700, 500), Margins = PdfCore.PageMargins.Uniform(20)
        }));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("Row001Row002", pdf.GetPage(1).Text);
        Assert.Equal("Row001Row003Row004", pdf.GetPage(2).Text);
        Assert.Equal("Row001Row003Row004Row005Row006Row007Row008", pdf.GetPage(3).Text);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, 2)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, 3)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, 2)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, 3)]
    public void SaveAsPdf_MiddleTitleRowsRepeatOnlyAfterTheirFirstOccurrence(ExcelPdfWorksheetLayoutMode layout, int firstBreak) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        for (int row = 1; row <= 8; row++) sheet.Cell(row, 1, "Row" + row.ToString("D3"));
        document.SetPrintArea(sheet, "A1:B8");
        document.SetPrintTitles(sheet, firstRow: 3, lastRow: 4, firstCol: null, lastCol: null);
        sheet.AddManualRowPageBreak(firstBreak); sheet.AddManualRowPageBreak(6);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PageSize(700, 500), Margins = PdfCore.PageMargins.Uniform(20)
        }));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.DoesNotContain("Row004", pdf.GetPage(1).Text);
        if (firstBreak == 2) Assert.DoesNotContain("Row003", pdf.GetPage(1).Text);
        else Assert.Contains("Row003", pdf.GetPage(1).Text);
        foreach (int page in new[] { 2, 3 }) {
            Assert.Contains("Row003", pdf.GetPage(page).Text);
            Assert.Contains("Row004", pdf.GetPage(page).Text);
            Assert.DoesNotContain("Row001", pdf.GetPage(page).Text);
        }
        Assert.Contains("Row005", pdf.GetPage(2).Text);
        Assert.DoesNotContain("Row007", pdf.GetPage(2).Text);
        Assert.Contains("Row007", pdf.GetPage(3).Text);
    }

    [Theory]
    [InlineData(true, 0, 6)]
    [InlineData(false, 0, 5)]
    [InlineData(true, 3, 3)]
    public void SaveAsPdf_CanvasRepeatsMiddleTitleRowsAcrossAutomaticPages(bool repeatTitles, int fitPages, int expectedPages) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        for (int row = 1; row <= 26; row++) {
            sheet.Cell(row, 1, "Row" + row.ToString("D3")); sheet.SetRowHeight(row, 30D);
        }
        document.SetPrintArea(sheet, "A1:A26");
        document.SetPrintTitles(sheet, firstRow: 3, lastRow: 4, firstCol: null, lastCol: null);
        if (fitPages > 0) sheet.SetPageSetup(fitToWidth: 0, fitToHeight: (uint)fitPages);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            UseWorksheetPrintTitleRows = repeatTitles,
            PageSize = new PdfCore.PageSize(300, 220), Margins = PdfCore.PageMargins.Uniform(20)
        }));
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        Assert.StartsWith("Row001", pdf.GetPage(1).Text);
        for (int page = 2; page <= expectedPages; page++) {
            if (repeatTitles) Assert.StartsWith("Row003Row004", pdf.GetPage(page).Text);
            else Assert.DoesNotContain("Row003", pdf.GetPage(page).Text);
            if (repeatTitles && fitPages == 0) Assert.Contains("Row" + (7 + (page - 2) * 4).ToString("D3"), pdf.GetPage(page).Text);
        }
        string text = string.Join(" ", pdf.GetPages().Select(page => page.Text));
        for (int row = 1; row <= 26; row++) Assert.Contains("Row" + row.ToString("D3"), text);
        Assert.Contains("Row026", pdf.GetPage(expectedPages).Text);
    }
}
