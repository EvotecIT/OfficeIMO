using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_PivotWithoutRenderedCellsIsCalculatedFromItsCache() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Sales");
        sheet.Cell(1, 1, "Region");
        sheet.Cell(1, 2, "Revenue");
        string[] regions = { "North", "South", "North", "East" };
        for (int index = 0; index < regions.Length; index++) {
            sheet.Cell(index + 2, 1, regions[index]);
            sheet.Cell(index + 2, 2, 100d * (index + 1));
        }

        sheet.AddPivotTable(
            sourceRange: "A1:B5",
            destinationCell: "D1",
            name: "RevenuePivot",
            rowFields: new[] { "Region" },
            dataFields: new[] { new ExcelPivotDataField("Revenue", ExcelPivotDataFunction.Sum, "Total Revenue") });

        using var stream = new MemoryStream();
        OfficeConversionReport report = document.SaveAsPdf(stream);
        using PdfPigDocument pdf = PdfPigDocument.Open(stream.ToArray());
        string text = string.Join(" ", Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).Select(word => word.Text));
        Assert.Contains("400", text.Split(' '));
        Assert.Contains("Total", text);
        // Rendering for the PDF must not write the pivot output into the caller's workbook.
        Assert.Contains("RevenuePivot", sheet.GetPivotTablesWithoutRenderedCells());
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code.Contains("pivot-table", StringComparison.Ordinal));
    }

    [Fact]
    public void BuiltInMediumTableStylesDrawHorizontalRulesOnly() {
        using ExcelDocument document = ExcelDocument.Create();
        Assert.True(ExcelBuiltInTableStylePaletteResolver.TryCreate(document, "TableStyleMedium2", out ExcelBuiltInTableStylePalette? medium));
        Assert.Equal(ExcelTableStyleBorderLayout.Horizontal, medium!.BorderLayout);
        Assert.True(ExcelBuiltInTableStylePaletteResolver.TryCreate(document, "TableStyleLight1", out ExcelBuiltInTableStylePalette? light));
        Assert.Equal(ExcelTableStyleBorderLayout.OutlineHorizontal, light!.BorderLayout);
        Assert.True(ExcelBuiltInTableStylePaletteResolver.TryCreate(document, "TableStyleLight15", out ExcelBuiltInTableStylePalette? grid));
        Assert.Equal(ExcelTableStyleBorderLayout.Grid, grid!.BorderLayout);
    }

    [Fact]
    public void SaveAsPdf_MediumTableStyleHasNoVerticalInnerLines() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Table");
        sheet.Cell(1, 1, "Name");
        sheet.Cell(1, 2, "Value");
        sheet.Cell(1, 3, "Other");
        for (int row = 2; row <= 4; row++) {
            sheet.Cell(row, 1, "Row " + row);
            sheet.Cell(row, 2, row);
            sheet.Cell(row, 3, row * 2);
        }

        sheet.AddTable("A1:C4", true, "Values", ExcelTableStyle.TableStyleMedium2);
        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        var page = pdf.GetPage(1);
        double nameX = FindWordStartX(page, "Name");
        double valueX = FindWordStartX(page, "Value");
        // A stroked vertical segment between the first two columns would be an inner vertical border.
        bool hasVerticalInnerLine = page.ExperimentalAccess.Paths
            .Select(path => path.GetBoundingRectangle())
            .Where(box => box.HasValue)
            .Select(box => box!.Value)
            .Any(box => box.Width < 1.5 && box.Height > 5 && box.Left > nameX + 5 && box.Right < valueX - 1);
        Assert.False(hasVerticalInnerLine);
    }

    [Fact]
    public void AutoFitColumns_AllowsForBoldTableHeaderAndFilterButton() {
        using ExcelDocument plain = ExcelDocument.Create();
        ExcelSheet plainSheet = plain.AddWorksheet("Plain");
        plainSheet.CellValue(1, 1, "Classification");
        plainSheet.CellValue(2, 1, "A");
        plainSheet.AutoFitColumns();

        using ExcelDocument table = ExcelDocument.Create();
        ExcelSheet tableSheet = table.AddWorksheet("Table");
        tableSheet.CellValue(1, 1, "Classification");
        tableSheet.CellValue(2, 1, "A");
        tableSheet.AddTable("A1:A2", true, "Classes", ExcelTableStyle.TableStyleMedium2);
        tableSheet.AutoFitColumns();

        double plainWidth = plainSheet.GetColumnDefinitions().Single(column => column.StartIndex <= 1 && column.EndIndex >= 1).Width!.Value;
        double tableWidth = tableSheet.GetColumnDefinitions().Single(column => column.StartIndex <= 1 && column.EndIndex >= 1).Width!.Value;
        // A filter button is about two and a half digit widths wide.
        Assert.True(tableWidth > plainWidth + 2D, $"Table header width {tableWidth} should exceed the plain width {plainWidth} by the filter button.");
    }

    [Fact]
    public void SaveAsPdf_WideChartPrintsAtFullScaleAcrossPages() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Chart");
        sheet.Cell(1, 1, "Quarter");
        sheet.Cell(1, 2, "Value");
        for (int row = 2; row <= 5; row++) {
            sheet.Cell(row, 1, "Q" + (row - 1));
            sheet.Cell(row, 2, row * 10);
        }

        sheet.AddChartFromRange("A1:B5", row: 7, column: 1, widthPixels: 1100, heightPixels: 300,
            type: ExcelChartType.ColumnClustered, title: "Wide chart");
        using PdfPigDocument pdf = OpenWorksheetPdf(document);
        // At 100 % the 825 pt chart does not fit a 468 pt Letter text width and continues on a second page.
        Assert.True(pdf.NumberOfPages >= 2);
        var quarter = pdf.GetPage(1).GetWords().First(word => word.Text == "Quarter");
        Assert.InRange(quarter.Letters[0].PointSize, 10.5, 11.5);
    }

    [Fact]
    public void SaveAsPdf_HeaderPictureStaysInsideTheTopMargin() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        sheet.Cell(1, 1, "FirstRow");
        sheet.SetHeaderFooter(headerCenter: "Logo Header");
        sheet.SetHeaderImage(ExcelHeaderFooterPosition.Center, CreateMinimalRgbPng(), "image/png", widthPoints: 96, heightPoints: 40);

        byte[] bytes = document.ToPdf(new ExcelPdfExportOptions {
            PageSize = new OfficeIMO.Pdf.PdfPageSize(420, 320),
            Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(54)
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

        OfficeIMO.Pdf.PdfImagePlacement image = Assert.Single(OfficeIMO.Pdf.PdfImageExtractor.ExtractImagePlacements(bytes));
        // PDF y grows upwards: the body starts 54 pt below the top of the 320 pt page.
        Assert.True(image.Y >= 320D - 54D - 0.01D, $"Header picture bottom {image.Y} overlaps the sheet body.");
        Assert.True(image.Y + image.Height <= 320D + 0.01D);
    }
}
