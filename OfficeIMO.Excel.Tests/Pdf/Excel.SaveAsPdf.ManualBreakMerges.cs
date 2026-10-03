using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, false)]
    public void SaveAsPdf_ManualBreakIsRetainedWhenNextMergeDoesNotCrossItsAxis(ExcelPdfWorksheetLayoutMode layout, bool rowBreak) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, "BeforeBreak");
        sheet.Cell(rowBreak ? 3 : 1, rowBreak ? 1 : 3, "AfterBreak");
        sheet.MergeRange(rowBreak ? "A3:B3" : "C1:C2");
        document.SetPrintArea(sheet, "A1:D4");
        if (rowBreak) {
            document.SetPrintTitles(sheet, firstRow: 3, lastRow: 3, firstCol: null, lastCol: null);
            sheet.AddManualRowPageBreak(2);
        } else {
            document.SetPrintTitles(sheet, firstRow: null, lastRow: null, firstCol: 3, lastCol: 3);
            sheet.AddManualColumnPageBreak(2);
        }
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PageSize(700, 500), Margins = PdfCore.PageMargins.Uniform(20)
        }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("BeforeBreak", pdf.GetPage(1).Text);
        Assert.DoesNotContain("AfterBreak", pdf.GetPage(1).Text);
        Assert.Contains("AfterBreak", pdf.GetPage(2).Text);
        Assert.DoesNotContain("BeforeBreak", pdf.GetPage(2).Text);
    }
}
