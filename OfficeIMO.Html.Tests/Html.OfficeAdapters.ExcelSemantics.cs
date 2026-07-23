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
        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);

        Assert.Contains("data-officeimo-value-kind=\"date-time\" data-officeimo-value=\"2026-07-11T14:15:16.0000000\"", html, StringComparison.Ordinal);
        Assert.True(importedSheet.TryGetCellValueSnapshot(2, 1, out ExcelCellValueSnapshot? snapshot));
        Assert.Equal(ExcelCellValueKind.DateTime, snapshot!.Kind);
        Assert.Equal(expected, snapshot.DateTimeValue);
        Assert.True(importedSheet.GetCellStyle(2, 1).IsDateLike);
        Assert.Empty(result.Report.Diagnostics);
    }

    [Fact]
    public void ExcelHtml_TableCellLimitRejectsOversizedSpanWithoutAllocation() {
        const string html = """
            <section class="officeimo-sheet" data-officeimo-sheet="Limited">
              <table><tr><td rowspan="50000" colspan="16000">Value</td></tr></table>
            </section>
            """;

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult(new HtmlToExcelOptions { MaxTableCells = 4 });
        using ExcelDocument workbook = result.Value;

        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
        Assert.Empty(Assert.Single(workbook.Sheets).GetMergedRanges());
    }

    [Fact]
    public void ExcelHtml_SemanticFormattingClampsHostileColumnSpanToNativeBounds() {
        const string html = """
            <table><tr><td colspan="2147483647"><strong>First</strong></td><td>Second</td></tr></table>
            """;

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, MaxTableCells = 4 });
        using ExcelDocument workbook = result.Value;

        ExcelSheet sheet = Assert.Single(workbook.Sheets);
        Assert.True(sheet.TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? first));
        Assert.Equal("First", first!.Text);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void ExcelHtml_TwoCellImageFallsBackBeforeEnumeratingAnOversizedAnchor() {
        const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/69DjmQAAAABJRU5ErkJggg==";
        string html = "<section class='officeimo-sheet' data-officeimo-sheet='Images'><table><tr><td>A</td></tr></table>"
            + "<section class='officeimo-images'><ul><li data-officeimo-anchor='twoCell' data-officeimo-row='1' data-officeimo-column='1' "
            + "data-officeimo-to-row='1048576' data-officeimo-to-column='16384'><span class='officeimo-feature-label'>Huge</span>"
            + "<img src='data:image/png;base64," + png + "'></li></ul></section></section>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { MaxTableCells = 4 });
        using ExcelDocument workbook = result.Value;
        ExcelImage image = Assert.Single(Assert.Single(workbook.Sheets).Images);

        Assert.False(image.HasTwoCellAnchor);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("two-cell anchor", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ExcelHtml_SvgIdNamespacingRewritesIdsAndReferencesInOnePass() {
        System.Reflection.MethodInfo method = typeof(ExcelHtmlConverterExtensions).GetMethod(
            "NamespaceSvgIds", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        const string svg = "<svg><defs><linearGradient id=\"g\"/><clipPath id='c'/></defs><rect fill=\"url(#g)\" clip-path=\"url(#c)\"/><use href=\"#g\" xlink:href='#c'/></svg>";

        string namespaced = (string)method.Invoke(null, new object[] { svg, "sheet-" })!;

        Assert.Contains("id=\"sheet-g\"", namespaced, StringComparison.Ordinal);
        Assert.Contains("id='sheet-c'", namespaced, StringComparison.Ordinal);
        Assert.Contains("url(#sheet-g)", namespaced, StringComparison.Ordinal);
        Assert.Contains("url(#sheet-c)", namespaced, StringComparison.Ordinal);
        Assert.Contains("href=\"#sheet-g\"", namespaced, StringComparison.Ordinal);
        Assert.Contains("xlink:href='#sheet-c'", namespaced, StringComparison.Ordinal);
    }
}
