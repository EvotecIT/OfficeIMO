using System.Globalization;
using AngleSharp.Html.Parser;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlExcelSemanticPresentation {
    private static ExcelDocument CreateReportWorkbook() {
        ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddSheet("Report");
        sheet.Cell(1, 1, "Title band");
        sheet.CellFontColor(1, 1, "FFFFFF");
        sheet.CellBackground(1, 1, "17365D");
        sheet.Cell(2, 1, "Amount"); sheet.Cell(2, 2, "Share"); sheet.Cell(2, 3, "Due"); sheet.Cell(2, 4, "Ready");
        sheet.Cell(3, 1, 1234567.5, numberFormat: "#,##0.00");
        sheet.Cell(3, 2, 0.3755186721991701, numberFormat: "0.0%");
        sheet.Cell(3, 3, new DateTime(2026, 10, 29), numberFormat: "yyyy-mm-dd");
        sheet.Cell(3, 4, true);
        sheet.Cell(4, 2, 0.9, numberFormat: "0.0%");
        sheet.AddConditionalColorScale("B3:B4", "F8D7DA", "C9ECD3");
        return workbook;
    }

    [Fact]
    public void SemanticTablesShowFillsNumberFormatsAndSpreadsheetAlignmentByDefault() {
        using ExcelDocument workbook = CreateReportWorkbook();
        string html = workbook.ToHtml();

        // The fill travels with the white text, so the title stays readable.
        Assert.Contains("background-color:#17365D", html, StringComparison.Ordinal);
        Assert.Contains("color:#FFFFFF", html, StringComparison.Ordinal);
        // Values use their number formats and keep the raw value for round trips.
        Assert.Contains(">1,234,567.50<", html, StringComparison.Ordinal);
        Assert.Contains(">37.6%<", html, StringComparison.Ordinal);
        Assert.Contains(">2026-10-29<", html, StringComparison.Ordinal);
        Assert.Contains(">TRUE<", html, StringComparison.Ordinal);
        var shareCell = Assert.Single(new HtmlParser().ParseDocument(html).QuerySelectorAll("td"), cell => cell.TextContent == "37.6%");
        Assert.Equal(0.3755186721991701, double.Parse(Assert.IsType<string>(shareCell.GetAttribute("data-officeimo-value")),
            NumberStyles.Float, CultureInfo.InvariantCulture));
        Assert.Contains("text-align:right", html, StringComparison.Ordinal);
        Assert.Contains("text-align:center", html, StringComparison.Ordinal);
        // The colour scale gives the two cells different fills.
        Assert.Contains("background-color:#F8D7DA", html, StringComparison.Ordinal);
        Assert.Contains("background-color:#C9ECD3", html, StringComparison.Ordinal);
        // Office fonts fall back to a sans-serif face instead of the browser's serif.
        Assert.DoesNotContain("font-family:&#39;Aptos&#39;\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void SemanticTablesCanKeepRawCellText() {
        using ExcelDocument workbook = CreateReportWorkbook();
        string html = workbook.ToHtml(new ExcelHtmlSaveOptions { ApplyCellPresentation = false });

        Assert.DoesNotContain("background-color:#17365D", html, StringComparison.Ordinal);
        Assert.DoesNotContain(">37.6%<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void SemanticTablesDrawChartsAndKeepDebugMetadataOptIn() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddSheet("Sales");
        sheet.Cell(1, 1, "Quarter"); sheet.Cell(1, 2, "Actual");
        sheet.Cell(2, 1, "Q1"); sheet.Cell(2, 2, 42);
        sheet.Cell(3, 1, "Q2"); sheet.Cell(3, 2, 58);
        sheet.CellFormula(4, 2, "SUM(B2:B3)");
        sheet.AddChart(OfficeChartType.ColumnClustered, "A1:B3", new ExcelChartOptions { Row = 6, Column = 1, Width = OfficeLength.Pixels(320), Height = OfficeLength.Pixels(200), Title = "Actual by quarter" });

        string html = workbook.ToHtml();

        Assert.Contains("<figure class=\"officeimo-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("<svg", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<details class=\"officeimo-chart-details\"><summary>Chart data</summary>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Type: ColumnClustered", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Series: 1; Categories: 2", html, StringComparison.Ordinal);
        Assert.DoesNotContain("officeimo-formulas", html, StringComparison.Ordinal);
    }
}
