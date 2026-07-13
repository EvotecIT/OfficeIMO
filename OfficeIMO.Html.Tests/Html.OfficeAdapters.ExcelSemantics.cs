using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdaptersExcelSemantics {
    [Fact]
    public void ExcelHtml_HeaderModeMakesHeaderIntentExplicit() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Not a header");
        sheet.CellValue(2, 1, "Second row");

        string defaultHtml = workbook.ToHtml();
        string dataOnlyHtml = workbook.ToHtml(new ExcelHtmlSaveOptions {
            HeaderMode = ExcelHtmlHeaderMode.None
        });

        Assert.Contains("<thead><tr><th data-officeimo-cell=\"A1\" scope=\"col\"", defaultHtml, StringComparison.Ordinal);
        Assert.Contains("<tbody><tr><td data-officeimo-cell=\"A1\"", dataOnlyHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("<thead>", dataOnlyHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_RoundTripsDateFormattedSerialAsDateTime() {
        DateTime expected = new(2026, 7, 11, 14, 15, 16, DateTimeKind.Unspecified);
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Dates");
        sheet.CellValue(1, 1, "When");
        sheet.CellValue(2, 1, expected);

        string html = workbook.ToHtml();
        HtmlToExcelResult result = html.ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);

        Assert.Contains("data-officeimo-value-kind=\"date-time\" data-officeimo-value=\"2026-07-11T14:15:16.0000000\"", html, StringComparison.Ordinal);
        Assert.True(importedSheet.TryGetCellValueSnapshot(2, 1, out ExcelCellValueSnapshot? snapshot));
        Assert.Equal(ExcelCellValueKind.DateTime, snapshot!.Kind);
        Assert.Equal(expected, snapshot.DateTimeValue);
        Assert.True(importedSheet.GetCellStyle(2, 1).IsDateLike);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void ExcelHtml_TableCellLimitRejectsOversizedSpanWithoutAllocation() {
        const string html = """
            <section class="officeimo-sheet" data-officeimo-sheet="Limited">
              <table><tr><td rowspan="50000" colspan="16000">Value</td></tr></table>
            </section>
            """;

        HtmlToExcelResult result = html.ToExcelDocumentResult(new HtmlToExcelOptions { MaxTableCells = 4 });
        using ExcelDocument workbook = result.Value;

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
        Assert.Empty(Assert.Single(workbook.Sheets).GetMergedRanges());
    }
}
