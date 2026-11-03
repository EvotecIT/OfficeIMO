using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, 2)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, 3)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, 2)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, 3)]
    public void SaveAsPdf_TitleBlockRepeatsOnlyColumnsBeforeTheCurrentBodySegment(ExcelPdfWorksheetLayoutMode layout, int firstBreak) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        for (int column = 1; column <= 8; column++) sheet.Cell(1, column).SetValue("Column" + column);
        sheet.PageSetup.PrintArea = "A1:H2";
        sheet.PageSetup.SetPrintTitles(firstRow: null, lastRow: null, firstColumn: 3, lastColumn: 4);
        sheet.AddManualColumnPageBreak(firstBreak); sheet.AddManualColumnPageBreak(6);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700),OfficeIMO.OfficeLength.Points(500)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.DoesNotContain("Column4", pdf.GetPage(1).Text);
        if (firstBreak == 2) Assert.DoesNotContain("Column3", pdf.GetPage(1).Text);
        else Assert.Contains("Column3", pdf.GetPage(1).Text);
        foreach (int page in new[] { 2, 3 }) {
            Assert.Contains("Column3", pdf.GetPage(page).Text);
            Assert.Contains("Column4", pdf.GetPage(page).Text);
            Assert.DoesNotContain("Column1", pdf.GetPage(page).Text);
        }
        Assert.Contains("Column5", pdf.GetPage(2).Text);
        Assert.DoesNotContain("Column7", pdf.GetPage(2).Text);
        Assert.Contains("Column7", pdf.GetPage(3).Text);
    }

    [Fact]
    public void ToPdfDocument_GeneralAlignmentUsesCellTypesAndRetainsExplicitAlignment() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1, 1).SetValue(123.45D); sheet.Cell(1, 2).SetValue("123.45"); sheet.Cell(1, 3).SetValue(true);
        sheet.Cell(1, 4).SetValue(123.45D); sheet.Cell(1, 4).SetHorizontalAlignment(ExcelHorizontalAlignment.Left);
        PdfCore.PdfDocument pdf = document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });
        PdfCore.TableBlock table = Assert.Single(Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdf.Blocks)).Blocks.OfType<PdfCore.TableBlock>());
        Assert.Equal(global::OfficeIMO.OfficeAlignment.Right, table.Style!.CellAlignments![(0, 0)]);
        Assert.False(table.Style.CellAlignments.ContainsKey((0, 1)));
        Assert.Equal(global::OfficeIMO.OfficeAlignment.Center, table.Style.CellAlignments[(0, 2)]);
        Assert.Equal("TRUE", table.Cells[0][2].Text);
        Assert.Equal(global::OfficeIMO.OfficeAlignment.Left, table.Style.CellAlignments[(0, 3)]);
    }

    [Fact]
    public void SaveAsPdf_CanvasUsesDefaultBottomAlignmentAndRetainsExplicitTopAlignment() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("DefaultBottom"); sheet.Cell(1, 2).SetValue("ExplicitTop");
        sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(40D)));
        sheet.Cell(1, 2).SetVerticalAlignment(ExcelVerticalAlignment.Top);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new ExcelPdfExportOptions { HeaderRowCount = 0 }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        ExcelSheet sheet = document.AddSheet("Report");
        for (int row = 1; row <= 80; row++) {
            sheet.Cell(row, 1).SetValue("Row" + row.ToString("D3") + (multiline ? "\nDetail" + row.ToString("D3") : ""));
            sheet.Cell(row, 1).SetFontSize(global::OfficeIMO.OfficeLength.Points(fontSize));
            sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(20D)));
        }
        sheet.PageSetup.FitToWidth = 1; sheet.PageSetup.FitToHeight = 2;
        byte[] bytes = document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(300)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        ExcelSheet sheet = document.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("Merged title");
        sheet.Range("A1:B1").Merge();
        sheet.Cell(1, 1).SetBold().SetFillColor("DDEEFF").SetFontColor("112233");
        sheet.Cell(1, 3).SetValue("First page"); sheet.Cell(1, 5).SetValue("Second page");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(10)); sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(20));
        sheet.PageSetup.SetPrintTitles(firstRow: null, lastRow: null, firstColumn: 1, lastColumn: 2);
        sheet.PageSetup.PrintArea = "A1:F2";
        sheet.AddManualColumnPageBreak(3);
        PdfCore.PdfDocument pdf = document.ToPdf(new ExcelPdfExportOptions {
            HeaderRowCount = 0, WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });
        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdf.Blocks));
        PdfCore.TableBlock[] tables = page.Blocks.OfType<PdfCore.TableBlock>().ToArray();
        Assert.Equal(2, tables.Length);
        foreach (PdfCore.TableBlock table in tables) {
            Assert.Equal(2, table.Cells[0][0].ColumnSpan);
            PdfCore.PdfTextRun run = Assert.Single(table.Cells[0][0].Runs);
            Assert.True(run.Bold);
            Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(17, 34, 51), run.Color);
            Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(221, 238, 255), table.Style!.CellFills![(0, 0)]);
            Assert.Equal(2D, table.Style.ColumnWidthWeights![1] / table.Style.ColumnWidthWeights[0]);
        }
    }

    [Theory]
    [InlineData(ExcelPageOrder.OverThenDown, false)]
    [InlineData(ExcelPageOrder.DownThenOver, false)]
    [InlineData(ExcelPageOrder.OverThenDown, true)]
    public void SaveAsPdf_CanvasHonorsAutomaticPageOrderAndFitPageCounts(ExcelPageOrder order, bool fit) {
        string path = Path.Combine(_directoryWithFiles, "AutomaticPageGrid-" + order + "-" + fit + ".xlsx");
        using ExcelDocument document = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Report" });
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1).SetValue("TopLeft"); sheet.Cell(1, 3).SetValue("TopRight");
        sheet.Cell(3, 1).SetValue("BottomLeft"); sheet.Cell(3, 3).SetValue("BottomRight");
        for (int column = 1; column <= 4; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(20));
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
        sheet.PageSetup.PageOrder = order; sheet.PageSetup.FitToWidth = fit ? 2U : null; sheet.PageSetup.FitToHeight = fit ? 2U : null;
        sheet.PageSetup.PrintArea = "A1:D4";
        byte[] bytes = document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(275),OfficeIMO.OfficeLength.Points(105)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
        using ExcelDocument document = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Report" });
        ExcelSheet sheet = document.Sheets[0];
        sheet.Cell(1, 1).SetValue("Corner");
        sheet.Cell(2, 1).SetValue("Row label");
        sheet.Cell(1, 3).SetValue("First heading");
        sheet.Cell(2, 3).SetValue("First value");
        sheet.Cell(1, 5).SetValue("Second heading");
        sheet.SetInternalLink(2, 5, "A2", display: "Back to row label");
        sheet.PageSetup.SetPrintTitles(firstRow: 1, lastRow: 1, firstColumn: 1, lastColumn: 1);
        sheet.AddManualColumnPageBreak(3);
        sheet.PageSetup.PrintArea = titlesOutsideArea ? "C2:F3" : "A1:F3";

        byte[] bytes = document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(700),OfficeIMO.OfficeLength.Points(500)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(20))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
