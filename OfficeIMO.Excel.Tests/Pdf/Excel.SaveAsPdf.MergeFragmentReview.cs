using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_MergeFragmentsAcceptAuthoredFractionalRowHeights(ExcelPdfWorksheetLayoutMode layout) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, "Start"); sheet.Cell(2, 1, "Merged"); sheet.Cell(4, 4, "End");
        sheet.MergeRange("A2:A3"); sheet.SetRowHeight(2, 15.1); sheet.SetRowHeight(3, 15.2);
        document.SetPrintArea(sheet, "A1:D4"); sheet.AddManualRowPageBreak(2);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PageSize(700, 500), Margins = PdfCore.PageMargins.Uniform(20)
        }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("Merged", pdf.GetPage(1).Text);
        Assert.Contains("Merged", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, true)]
    public void SaveAsPdf_ImageBearingMergeRetainsItsImageAcrossPageBreaks(ExcelPdfWorksheetLayoutMode layout, bool horizontal) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, 36);
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, 15);
        sheet.Cell(1, 1, "Start"); sheet.Cell(2, 1, "Merged"); sheet.Cell(4, 4, "End");
        sheet.MergeRange(horizontal ? "A2:B2" : "A2:A3");
        sheet.AddImage(2, 1, CreateMinimalRgbPng(), "image/png", widthPixels: 24, heightPixels: 16);
        document.SetPrintArea(sheet, "A1:D4");
        if (horizontal) sheet.AddManualColumnPageBreak(1); else sheet.AddManualRowPageBreak(2);
        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PageSize(700, 500), Margins = PdfCore.PageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Contains("Merged", string.Concat(pdf.GetPages().Select(page => page.Text)));
    }
}
