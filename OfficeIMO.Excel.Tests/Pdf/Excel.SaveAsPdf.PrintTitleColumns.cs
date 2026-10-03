using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void ToPdfDocument_GeneralAlignmentUsesCellTypesAndRetainsExplicitAlignment() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, 123.45D); sheet.Cell(1, 2, "123.45"); sheet.Cell(1, 3, true);
        sheet.Cell(1, 4, 123.45D); sheet.CellAlign(1, 4, ExcelHorizontalAlignment.Left);
        PdfCore.PdfDocument pdf = document.ToPdfDocument(new ExcelToPdfOptions {
            HeaderRowCount = 0, WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });
        PdfCore.TableBlock table = Assert.Single(Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdf.Blocks)).Blocks.OfType<PdfCore.TableBlock>());
        Assert.Equal(PdfCore.PdfColumnAlign.Right, table.Style!.CellAlignments![(0, 0)]);
        Assert.False(table.Style.CellAlignments.ContainsKey((0, 1)));
        Assert.Equal(PdfCore.PdfColumnAlign.Center, table.Style.CellAlignments[(0, 2)]);
        Assert.Equal("TRUE", table.Cells[0][2].Text);
        Assert.Equal(PdfCore.PdfColumnAlign.Left, table.Style.CellAlignments[(0, 3)]);
    }

    [Fact]
    public void SaveAsPdf_CanvasUsesDefaultBottomAlignmentAndRetainsExplicitTopAlignment() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, "DefaultBottom"); sheet.Cell(1, 2, "ExplicitTop");
        sheet.SetRowHeight(1, 40D);
        sheet.CellVerticalAlign(1, 2, ExcelVerticalAlignment.Top);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new ExcelToPdfOptions { HeaderRowCount = 0 }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        var bottom = Assert.Single(words, word => word.Text == "DefaultBottom");
        var top = Assert.Single(words, word => word.Text == "ExplicitTop");
        Assert.InRange(top.BoundingBox.Bottom - bottom.BoundingBox.Bottom, 25D, 35D);
    }

    [Theory]
    [InlineData(11D, false)]
    [InlineData(20D, true)]
    public void SaveAsPdf_FlowFitToHeightScalesExplicitCellFontsWithoutLosingRows(double fontSize, bool multiline) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        for (int row = 1; row <= 80; row++) {
            sheet.Cell(row, 1, "Row" + row.ToString("D3") + (multiline ? "\nDetail" + row.ToString("D3") : ""));
            sheet.CellFontSize(row, 1, fontSize);
            sheet.SetRowHeight(row, 20D);
        }
        sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 2);
        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            HeaderRowCount = 0, WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable,
            PageSize = new PdfCore.PageSize(300, 300), Margins = PdfCore.PageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        string text = string.Join(" ", pdf.GetPages().Select(page => page.Text));
        for (int row = 1; row <= 80; row++) {
            Assert.Contains("Row" + row.ToString("D3"), text, StringComparison.Ordinal);
            if (multiline) Assert.Contains("Detail" + row.ToString("D3"), text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void ToPdfDocument_RepeatedColumnsPreserveMergedTitleFormattingAndWidths() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, "Merged title");
        sheet.Range("A1:B1").Merge();
        sheet.CellAt(1, 1).SetBold().SetFillColor("DDEEFF").SetFontColor("112233");
        sheet.Cell(1, 3, "First page"); sheet.Cell(1, 5, "Second page");
        sheet.SetColumnWidth(1, 10); sheet.SetColumnWidth(2, 20);
        document.SetPrintTitles(sheet, firstRow: null, lastRow: null, firstCol: 1, lastCol: 2);
        document.SetPrintArea(sheet, "A1:F2");
        sheet.AddManualColumnPageBreak(3);
        PdfCore.PdfDocument pdf = document.ToPdfDocument(new ExcelToPdfOptions {
            HeaderRowCount = 0, WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });
        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdf.Blocks));
        PdfCore.TableBlock[] tables = page.Blocks.OfType<PdfCore.TableBlock>().ToArray();
        Assert.Equal(2, tables.Length);
        foreach (PdfCore.TableBlock table in tables) {
            Assert.Equal(2, table.Cells[0][0].ColumnSpan);
            PdfCore.PdfTextRun run = Assert.Single(table.Cells[0][0].Runs);
            Assert.True(run.Bold);
            Assert.Equal(PdfCore.PdfColor.FromRgb(17, 34, 51), run.Color);
            Assert.Equal(PdfCore.PdfColor.FromRgb(221, 238, 255), table.Style!.CellFills![(0, 0)]);
            Assert.Equal(2D, table.Style.ColumnWidthWeights![1] / table.Style.ColumnWidthWeights[0]);
        }
    }

    [Theory]
    [InlineData(ExcelPageOrder.OverThenDown, false)]
    [InlineData(ExcelPageOrder.DownThenOver, false)]
    [InlineData(ExcelPageOrder.OverThenDown, true)]
    public void SaveAsPdf_CanvasHonorsAutomaticPageOrderAndFitPageCounts(ExcelPageOrder order, bool fit) {
        string path = Path.Combine(_directoryWithFiles, "AutomaticPageGrid-" + order + "-" + fit + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, "Report");
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1, "TopLeft"); sheet.Cell(1, 3, "TopRight");
        sheet.Cell(3, 1, "BottomLeft"); sheet.Cell(3, 3, "BottomRight");
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, 20);
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, 30);
        sheet.SetPageSetup(pageOrder: order, fitToWidth: fit ? 2U : null, fitToHeight: fit ? 2U : null);
        document.SetPrintArea(sheet, "A1:D4");
        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(275, 105), Margins = PdfCore.PageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(4, pdf.NumberOfPages);
        string[] markers = order == ExcelPageOrder.OverThenDown
            ? new[] { "TopLeft", "TopRight", "BottomLeft", "BottomRight" }
            : new[] { "TopLeft", "BottomLeft", "TopRight", "BottomRight" };
        for (int index = 0; index < markers.Length; index++) Assert.Contains(markers[index], pdf.GetPage(index + 1).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, false)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, true)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, true)]
    public void SaveAsPdf_RepeatsTitleColumnsAcrossManualColumnPages(ExcelPdfWorksheetLayoutMode layout, bool titlesOutsideArea) {
        string path = Path.Combine(_directoryWithFiles, "RepeatedColumns-" + layout + "-" + titlesOutsideArea + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, "Report");
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1, "Corner");
        sheet.Cell(2, 1, "Row label");
        sheet.Cell(1, 3, "First heading");
        sheet.Cell(2, 3, "First value");
        sheet.Cell(1, 5, "Second heading");
        sheet.SetInternalLink(2, 5, "A2", display: "Back to row label");
        document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1);
        sheet.AddManualColumnPageBreak(3);
        document.SetPrintArea(sheet, titlesOutsideArea ? "C2:F3" : "A1:F3");

        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PageSize(700, 500), Margins = PdfCore.PageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (var page in pdf.GetPages()) {
            Assert.Contains("Corner", page.Text, StringComparison.Ordinal);
            Assert.Contains("Row label", page.Text, StringComparison.Ordinal);
        }
        Assert.Contains("First value", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("First value", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("Back to row label", pdf.GetPage(2).Text, StringComparison.Ordinal);
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(bytes);
        PdfCore.PdfNamedDestination target = Assert.Single(logical.NamedDestinations, item => item.Name.EndsWith("-a2", StringComparison.Ordinal));
        Assert.Equal(1, target.PageNumber);
        Assert.Single(logical.GetLinksByDestinationName(target.Name));
        Assert.True(document.InspectFeatures().Can(ExcelPreflightCapability.ExportPdfReport));
    }
}
