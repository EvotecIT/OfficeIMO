using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdaptersExcelMergedCells {
    [Fact]
    public void ExcelHtml_DefaultSemanticExportLimitsAreBounded() {
        var options = new ExcelHtmlSaveOptions();

        Assert.Equal(ExcelHtmlSaveOptions.DefaultMaxRowsPerSheet, options.MaxRowsPerSheet);
        Assert.Equal(ExcelHtmlSaveOptions.DefaultMaxColumnsPerSheet, options.MaxColumnsPerSheet);
        Assert.Equal(ExcelHtmlSaveOptions.DefaultMaxCellsPerSheet, options.MaxCellsPerSheet);
        Assert.Equal(ExcelHtmlSaveOptions.DefaultMaxMergedRangesPerSheet, options.MaxMergedRangesPerSheet);
    }

    [Fact]
    public void ExcelHtml_RoundTripsMergedCellsWithoutDuplicatingCoveredCells() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Merged");
        sheet.CellValue(1, 1, "Quarterly result");
        sheet.CellValue(1, 4, "Status");
        sheet.CellValue(4, 2, "Approved");
        sheet.MergeRange("A1:C2");
        sheet.MergeRange("B4:C4");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables
        });

        Assert.Contains("data-officeimo-cell=\"A1\" rowspan=\"2\" colspan=\"3\" data-officeimo-merge=\"A1:C2\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-cell=\"B4\" colspan=\"2\" data-officeimo-merge=\"B4:C4\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-officeimo-cell=\"B1\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-officeimo-cell=\"A2\"", html, StringComparison.Ordinal);

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;
        ExcelSheet importedSheet = Assert.Single(imported.Sheets);
        string[] mergedRanges = importedSheet.GetMergedRanges().Select(merge => merge.A1Range).OrderBy(range => range).ToArray();

        Assert.Equal(2, result.MergedRanges);
        Assert.Equal(new[] { "A1:C2", "B4:C4" }, mergedRanges);
        Assert.True(importedSheet.TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? title));
        Assert.Equal("Quarterly result", title!.Text);
        Assert.True(importedSheet.TryGetCellValueSnapshot(4, 2, out ExcelCellValueSnapshot? approval));
        Assert.Equal("Approved", approval!.Text);
    }

    [Fact]
    public void ExcelHtml_ImportsGenericTableRowAndColumnSpans() {
        const string html = """
            <main>
              <section class="officeimo-sheet" data-officeimo-sheet="Generic" data-officeimo-range="B2:E4">
                <table>
                  <tr><th rowspan="2" colspan="2">Group</th><th>Q1</th></tr>
                  <tr><td>10</td></tr>
                  <tr><td>Tail</td><td colspan="2">Summary</td></tr>
                </table>
              </section>
            </main>
            """;

        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument workbook = result.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);

        Assert.Equal(2, result.MergedRanges);
        Assert.Equal(new[] { "B2:C3", "C4:D4" }, sheet.GetMergedRanges().Select(merge => merge.A1Range).OrderBy(range => range).ToArray());
        Assert.True(sheet.TryGetCellValueSnapshot(2, 4, out ExcelCellValueSnapshot? q1));
        Assert.Equal("Q1", q1!.Text);
        Assert.True(sheet.TryGetCellValueSnapshot(3, 4, out ExcelCellValueSnapshot? ten));
        Assert.Equal("10", ten!.Text);
    }

    [Fact]
    public void ExcelHtml_TruncationClipsMergedRangeToExportedRows() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Clipped");
        sheet.CellValue(1, 1, "Visible");
        sheet.MergeRange("A1:B3");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            Profile = OfficeHtmlConversionProfile.ExcelSemanticTables,
            MaxRowsPerSheet = 2
        });
        HtmlToExcelResult result = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using ExcelDocument imported = result.Value;

        Assert.Contains("data-officeimo-merge=\"A1:B2\"", html, StringComparison.Ordinal);
        Assert.Equal("A1:B2", Assert.Single(Assert.Single(imported.Sheets).GetMergedRanges()).A1Range);
    }

    [Fact]
    public void ExcelHtml_HugeMergeIsClippedBeforeCellMaterialization() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("Bounded");
        sheet.CellValue(1, 1, "Visible");
        sheet.MergeRange("A1:XFD1048576");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            MaxRowsPerSheet = 2,
            MaxColumnsPerSheet = 3,
            MaxCellsPerSheet = 6
        });

        Assert.Contains("data-officeimo-merge=\"A1:C2\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-officeimo-cell=\"D1\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-officeimo-cell=\"A3\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_EmptyHugeMergeUsesTheBoundedProbeWindow() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("EmptyBounded");
        sheet.MergeRange("A1:XFD1048576");

        string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
            MaxRowsPerSheet = 2,
            MaxColumnsPerSheet = 3,
            MaxCellsPerSheet = 4
        });

        Assert.Contains("data-officeimo-merge=\"A1:C1\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-officeimo-cell=\"A2\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelHtml_RejectsMergeMetadataBeyondTheConfiguredLimit() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = workbook.AddWorksheet("MergeBounded");
        sheet.MergeRange("A1:B1");
        sheet.MergeRange("A2:B2");

        Assert.Throws<InvalidOperationException>(() => workbook.ToHtml(new ExcelHtmlSaveOptions {
            MaxMergedRangesPerSheet = 1
        }));
    }
}
