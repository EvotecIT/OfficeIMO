using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    private static PdfCore.PdfTableStyle ExportFlowTableStyle(ExcelDocument document) {
        PdfCore.PdfDocument pdf = document.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 0,
            WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable
        });
        PdfCore.PageBlock page = Assert.IsType<PdfCore.PageBlock>(Assert.Single(pdf.Blocks));
        PdfCore.TableBlock table = Assert.Single(page.Blocks.OfType<PdfCore.TableBlock>());
        return Assert.IsType<PdfCore.PdfTableStyle>(table.Style);
    }

    [Fact]
    public void SaveAsPdf_PrintsRuleBasedConditionalFills() {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "ExcelPdfRuleConditionalFills.xlsx"));
        ExcelSheet sheet = document.AddSheet("Rota");
        sheet.Cell(1, 1, "Maya"); sheet.Cell(1, 2, "D");
        sheet.Cell(2, 1, "Leon"); sheet.Cell(2, 2, "R");
        sheet.Range("B1:B2").ConditionalFormatting.ContainsText("D", "DCFCE7");

        PdfCore.PdfTableStyle style = ExportFlowTableStyle(document);

        Assert.NotNull(style.CellFills);
        Assert.Equal(OfficeColor.FromRgb(0xDC, 0xFC, 0xE7), style.CellFills![(0, 1)]);
        Assert.False(style.CellFills.ContainsKey((1, 1)));
    }

    [Fact]
    public void SaveAsPdf_PrintsPivotTablesWithTheirPivotStyle() {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "ExcelPdfPivotStyle.xlsx"));
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Sales");
        sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, 10d);
        sheet.CellValue(3, 1, "South"); sheet.CellValue(3, 2, 12d);
        sheet.AddPivotTable(sourceRange: "A1:B3", destinationCell: "D1", new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = new[] { "Region" }, DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales") }, PivotStyleName = "PivotStyleMedium9" });

        PdfCore.PdfTableStyle style = ExportFlowTableStyle(document);

        // The pivot's header row takes the style's header fill; the source table has none.
        Assert.NotNull(style.CellFills);
        Assert.True(style.CellFills!.ContainsKey((0, 3)));
        Assert.False(style.CellFills.ContainsKey((0, 0)));
    }

    [Fact]
    public void SaveAsPdf_SetsHeaderPicturesInlineOnTheHeaderLine() {
        string path = Path.Combine(_directoryWithFiles, "ExcelPdfInlineHeaderPicture.xlsx");
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Report" })) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Body");
            sheet.SetHeaderFooter(headerCenter: "Scan&G");
            sheet.SetHeaderImage(ExcelHeaderFooterPosition.Center, CreateMinimalRgbPng(), width: OfficeLength.Points(24), height: OfficeLength.Points(18));
            document.Save();
            bytes = document.ToPdf(new ExcelPdfExportOptions { IncludeSheetHeadings = false })
                .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var image = Assert.Single(page.GetImages());
        double baseline = FindWordStartY(page, "Scan");
        // The picture stands on the header baseline beside the text instead of hanging below it.
        Assert.InRange(image.Bounds.Bottom, baseline - 1D, baseline + 1D);
        Assert.True(image.Bounds.Top <= page.Height - 0.3D * 72D + 1D, "The picture starts below the 0.3 inch header margin.");
        Assert.True(image.Bounds.Bottom >= page.Height - 0.75D * 72D, "The picture ends above the sheet.");
    }

    [Fact]
    public void SaveAsPdf_KeepsShortCategoryLabelsThatFitTheirSlots() {
        using ExcelDocument document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "ExcelPdfCategoryLabels.xlsx"));
        ExcelSheet sheet = document.AddSheet("Chart");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Revenue");
        string[] regions = { "North", "South", "East", "West" };
        for (int index = 0; index < regions.Length; index++) {
            sheet.CellValue(index + 2, 1, regions[index]);
            sheet.CellValue(index + 2, 2, 100d + index * 10d);
        }

        sheet.AddChart(OfficeChartType.ColumnClustered, "A1:B5", new ExcelChartOptions { Row = 7, Column = 1, Width = OfficeLength.Pixels(300), Height = OfficeLength.Pixels(200), Title = "Revenue" });
        byte[] bytes = document.ToPdf(new ExcelPdfExportOptions { IncludeSheetHeadings = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(Enumerable.Range(1, pdf.NumberOfPages).Select(index => pdf.GetPage(index).Text));
        foreach (string region in regions) Assert.True(text.Split(new[] { region }, StringSplitOptions.None).Length >= 3, region + " should appear in the cells and on the axis.");
    }
}
