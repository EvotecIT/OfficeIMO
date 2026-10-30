using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
    public void SaveAsPdf_ImportedFractionalRowHeightsRetainMergedTextAcrossManualBreak(ExcelPdfWorksheetLayoutMode layout) {
        using var workbook = new MemoryStream();
        using (SpreadsheetDocument package = SpreadsheetDocument.Create(workbook, SpreadsheetDocumentType.Workbook, true)) {
            WorkbookPart workbookPart = package.AddWorkbookPart();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("sheet");
            workbookPart.Workbook = new Workbook(
                new Sheets(new Sheet { Id = "sheet", SheetId = 1U, Name = "Report" }),
                new DefinedNames(new DefinedName {
                    Name = "_xlnm.Print_Area", LocalSheetId = 0U, Text = "'Report'!$A$1:$D$4"
                }));
            // Import unquantized OOXML heights: each fragment rounds up in EMUs,
            // while the full two-row height rounds down. SetRowHeight would hide it.
            const string worksheetXml = """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <dimension ref="A1:D4"/>
                  <sheetFormatPr defaultRowHeight="36"/>
                  <cols><col min="1" max="4" width="15" customWidth="1"/></cols>
                  <sheetData>
                    <row r="1"><c r="A1" t="inlineStr"><is><t>FirstPage</t></is></c></row>
                    <row r="2" ht="15.1236" customHeight="1"><c r="A2" t="inlineStr"><is><t>Merged</t></is></c></row>
                    <row r="3" ht="15.1236" customHeight="1"/>
                    <row r="4"><c r="D4" t="inlineStr"><is><t>LastPage</t></is></c></row>
                  </sheetData>
                  <mergeCells count="1"><mergeCell ref="A2:A3"/></mergeCells>
                  <rowBreaks count="1" manualBreakCount="1"><brk id="2" min="0" max="16383" man="1"/></rowBreaks>
                </worksheet>
                """;
            using var worksheetXmlStream = new MemoryStream(Encoding.UTF8.GetBytes(worksheetXml));
            worksheetPart.FeedData(worksheetXmlStream);
        }

        workbook.Position = 0;
        using ExcelDocument document = ExcelDocument.Load(workbook);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700), OfficeIMO.OfficeLength.Points(500)),
            Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.False(string.IsNullOrWhiteSpace(pdf.GetPage(1).Text));
        Assert.False(string.IsNullOrWhiteSpace(pdf.GetPage(2).Text));
        Assert.Contains("FirstPage", pdf.GetPage(1).Text);
        Assert.Contains("LastPage", pdf.GetPage(2).Text);
        Assert.Contains("Merged", pdf.GetPage(2).Text);
        Assert.DoesNotContain("Merged", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, true)]
    public void SaveAsPdf_ConditionalMergeDecorationsRetainTheirFullGeometryAcrossBreaks(ExcelPdfWorksheetLayoutMode layout, bool bar) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(36)));
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(15));
        sheet.Cell(1, 1, 0); sheet.Cell(2, 2, 50); sheet.Cell(4, 4, 100);
        sheet.MergeRange("B2:C3");
        if (bar) sheet.AddConditionalDataBar("A1:D4", "FF5B9BD5");
        else sheet.AddConditionalColorScale("A1:D4", "FFFF0000", "FF00FF00");
        document.SetPrintArea(sheet, "A1:D4");
        sheet.AddManualRowPageBreak(2); sheet.AddManualColumnPageBreak(2);
        byte[] bytes = document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700),OfficeIMO.OfficeLength.Points(500)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("50", pdf.GetPage(4).Text);
        string content = PdfOperatorSearchText.From(bytes);
        string color = bar ? "0.357 0.608 0.835 rg" : "0.502 0.502 0 rg";
        Assert.Equal(bar ? 3 : 4, content.Split(new[] { color }, StringSplitOptions.None).Length - 1);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_MergedDiagonalBordersRemainOnTheFragmentContainingTheSourceAnchor(ExcelPdfWorksheetLayoutMode layout) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(36)));
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(15));
        sheet.Cell(1, 1, "Start"); sheet.Cell(2, 2, "End");
        sheet.MergeRange("B2:C3");
        sheet.CellDiagonalBorder(2, 2, ExcelBorderStyle.Thin, "B91C1C", diagonalUp: true, diagonalDown: true);
        document.SetPrintArea(sheet, "A1:D4");
        sheet.AddManualRowPageBreak(2); sheet.AddManualColumnPageBreak(2);
        byte[] bytes = document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700),OfficeIMO.OfficeLength.Points(500)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(4, pdf.NumberOfPages);
        string content = PdfOperatorSearchText.From(bytes);
        Assert.Equal(2, content.Split(new[] { "0.725 0.11 0.11 RG" }, StringSplitOptions.None).Length - 1);
    }
    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_MergedCenteredWrappedTextPreservesFullyVisibleLinesOnBothSidesOfTheBreak(ExcelPdfWorksheetLayoutMode layout) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(50)));
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(15));
        sheet.Cell(1, 1, "Start"); sheet.Cell(4, 4, "End");
        sheet.MergeRange("A2:B3");
        sheet.Cell(2, 1, "WrapOne\nWrapTwo\nWrapThree\nWrapFour");
        sheet.CellAlign(2, 1, ExcelHorizontalAlignment.Center);
        sheet.CellVerticalAlign(2, 1, ExcelVerticalAlignment.Center);
        sheet.CellFontName(2, 1, "Arial"); sheet.CellFontSize(2, 1, global::OfficeIMO.OfficeLength.Points(11)); sheet.CellWrapText(2, 1);
        document.SetPrintArea(sheet, "A1:D4"); sheet.AddManualRowPageBreak(2);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700),OfficeIMO.OfficeLength.Points(500)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("WrapOne", pdf.GetPage(1).Text); Assert.Contains("WrapTwo", pdf.GetPage(1).Text);
        // A glyph intersecting the page cut is clipped visually. PDF text readers
        // can still report that partial line, so assert only wholly excluded lines.
        Assert.DoesNotContain("WrapFour", pdf.GetPage(1).Text);
        Assert.Contains("WrapThree", pdf.GetPage(2).Text); Assert.Contains("WrapFour", pdf.GetPage(2).Text);
        Assert.DoesNotContain("WrapOne", pdf.GetPage(2).Text);
    }
    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_MergeCrossingBothBreakAxesKeepsRightBottomTextOnTheLastPage(ExcelPdfWorksheetLayoutMode layout) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(36)));
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(15));
        sheet.Cell(1, 1, "Start");
        sheet.Cell(2, 2, "End");
        sheet.CellAlign(2, 2, ExcelHorizontalAlignment.Right);
        sheet.CellVerticalAlign(2, 2, ExcelVerticalAlignment.Bottom);
        sheet.MergeRange("B2:C3");
        document.SetPrintArea(sheet, "A1:D4");
        sheet.AddManualRowPageBreak(2);
        sheet.AddManualColumnPageBreak(2);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700),OfficeIMO.OfficeLength.Points(500)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("Start", pdf.GetPage(1).Text);
        Assert.Contains("End", pdf.GetPage(4).Text);
        for (int page = 1; page <= 3; page++) Assert.DoesNotContain("End", pdf.GetPage(page).Text);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, true, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, false, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, true, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, false, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, true, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, true, false)]
    public void SaveAsPdf_ManualBreakAcrossMergeRetainsTheFullCellTextGeometry(ExcelPdfWorksheetLayoutMode layout, bool rowBreak, bool explicitAlignment) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(36)));
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(15));
        sheet.Cell(1, 1, "FirstPage");
        sheet.Cell(4, 4, "LastPage");
        sheet.Cell(rowBreak ? 2 : 1, rowBreak ? 1 : 2, "Merged");
        if (explicitAlignment) sheet.CellVerticalAlign(rowBreak ? 2 : 1, rowBreak ? 1 : 2, ExcelVerticalAlignment.Bottom);
        sheet.MergeRange(rowBreak ? "A2:A3" : "B1:C1");
        document.SetPrintArea(sheet, "A1:D4");
        if (rowBreak) sheet.AddManualRowPageBreak(2);
        else sheet.AddManualColumnPageBreak(2);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700),OfficeIMO.OfficeLength.Points(500)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FirstPage", pdf.GetPage(1).Text);
        Assert.Contains("LastPage", pdf.GetPage(2).Text);
        Assert.Contains("Merged", pdf.GetPage(rowBreak ? 2 : 1).Text);
        Assert.DoesNotContain("Merged", pdf.GetPage(rowBreak ? 1 : 2).Text);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, false)]
    public void SaveAsPdf_ManualBreakIsRetainedWhenNextMergeDoesNotCrossItsAxis(ExcelPdfWorksheetLayoutMode layout, bool rowBreak) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
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
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700),OfficeIMO.OfficeLength.Points(500)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("BeforeBreak", pdf.GetPage(1).Text);
        Assert.DoesNotContain("AfterBreak", pdf.GetPage(1).Text);
        Assert.Contains("AfterBreak", pdf.GetPage(2).Text);
        Assert.DoesNotContain("BeforeBreak", pdf.GetPage(2).Text);
    }
}
