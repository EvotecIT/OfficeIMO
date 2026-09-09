using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlExcelReportValues {
    [Theory]
    [InlineData("First\n", "Second")]
    [InlineData(" First", "  Second ")]
    [InlineData("First\t", "Second")]
    public void SemanticRichTextPreservesWhitespaceAndRunFormatting(string first, string second) {
        using ExcelDocument source = ExcelDocument.Create(new MemoryStream());
        source.AddWorksheet("Notes").CellAt(1, 1).SetRichText(
            new ExcelRichTextRun(first) { Bold = true },
            new ExcelRichTextRun(second) { Italic = true });
        string html = source.ToHtml(new ExcelHtmlSaveOptions { HeaderMode = ExcelHtmlHeaderMode.None });
        using ExcelDocument imported = HtmlConversionDocument.Parse(html).ToExcelDocument();
        using MemoryStream artifact = imported.ToStream();
        using ExcelDocument reopened = ExcelDocument.Load(artifact);
        ExcelCell cell = Assert.Single(reopened.Sheets).CellAt(1, 1);
        Assert.Equal(first + second, cell.GetValue<string>());
        Assert.Contains(cell.GetRichText(), run => run.Text == first && run.Bold);
        Assert.Contains(cell.GetRichText(), run => run.Text == second && run.Italic);
    }

    [Theory]
    [InlineData(HtmlImportMode.Semantic)]
    [InlineData(HtmlImportMode.Generic)]
    public void AnnotatedTextCannotBeReplacedByDifferentRichText(HtmlImportMode mode) {
        const string html = "<section class='officeimo-sheet' data-officeimo-sheet='Notes'><table><tr>"
            + "<td style='white-space:pre' data-officeimo-value-kind='text' data-officeimo-value='A B'><strong>A\nB</strong></td>"
            + "</tr></table></section>";
        using ExcelDocument workbook = HtmlConversionDocument.Parse(html).ToExcelDocument(
            new HtmlToExcelOptions { Mode = mode, ImportTypedCellValues = true });
        Assert.Equal("A B", Assert.Single(workbook.Sheets).CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(HtmlImportMode.Generic)]
    [InlineData(HtmlImportMode.Auto)]
    public void ExplicitReportValuesRemainTypedAfterSaveAndReopen(HtmlImportMode mode) {
        const string html = """
            <table><caption>Service usage</caption>
              <thead><tr><th>Reference</th><th>Hours</th><th>Approved</th><th>Recorded</th></tr></thead>
              <tbody><tr>
                <td>00127</td>
                <td style="color: #123456" data-officeimo-value-kind="number" data-officeimo-value="12.5"><strong style="color: #a12030">12.50</strong></td>
                <td data-officeimo-value-kind="boolean" data-officeimo-value="true"><em>Yes</em></td>
                <td data-officeimo-value-kind="date-time" data-officeimo-value="2026-09-01T00:00:00">1 September 2026</td>
              </tr></tbody>
            </table>
            """;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = mode, ImportTypedCellValues = true });
        using ExcelDocument workbook = result.RequireValue();
        using MemoryStream artifact = workbook.ToStream();
        using ExcelDocument reopened = ExcelDocument.Load(artifact);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);

        Assert.Equal("Service usage", sheet.Name);
        Assert.Equal(ExcelCellValueKind.Text, Snapshot(sheet, 2, 1).Kind);
        Assert.Equal("00127", Snapshot(sheet, 2, 1).Text);
        Assert.Equal(ExcelCellValueKind.Number, Snapshot(sheet, 2, 2).Kind);
        Assert.Equal(12.5D, sheet.CellAt(2, 2).GetValue<double>());
        Assert.True(sheet.GetCellStyle(2, 2).Bold);
        Assert.Equal("FFA12030", sheet.GetCellStyle(2, 2).FontColorArgb);
        Assert.Equal(ExcelCellValueKind.Boolean, Snapshot(sheet, 2, 3).Kind);
        Assert.True(sheet.CellAt(2, 3).GetValue<bool>());
        Assert.Equal(ExcelCellValueKind.DateTime, Snapshot(sheet, 2, 4).Kind);
        Assert.Equal(new DateTime(2026, 9, 1), Snapshot(sheet, 2, 4).DateTimeValue);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TypedValuesAreExplicitAndDoNotGuessLiteralIdentifiers(bool importTypedValues) {
        const string html = """
            <table><tr>
              <td data-officeimo-value-kind="text" data-officeimo-value="00027"><strong>Account 00027</strong></td>
              <td data-officeimo-value-kind="number" data-officeimo-value="12.5">12,50 hours</td>
              <td>00127</td><td>1E12</td><td>=1+1</td>
            </tr></table>
            """;
        using ExcelDocument workbook = HtmlConversionDocument.Parse(html).ToExcelDocument(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, ImportTypedCellValues = importTypedValues });
        ExcelSheet sheet = Assert.Single(workbook.Sheets);

        Assert.Equal(importTypedValues ? "00027" : "Account 00027", Snapshot(sheet, 1, 1).Text);
        Assert.Equal(importTypedValues ? ExcelCellValueKind.Number : ExcelCellValueKind.Text, Snapshot(sheet, 1, 2).Kind);
        Assert.Equal("00127", Snapshot(sheet, 1, 3).Text);
        Assert.Equal("1E12", Snapshot(sheet, 1, 4).Text);
        Assert.Equal(ExcelCellValueKind.Text, Snapshot(sheet, 1, 4).Kind);
        Assert.Equal("=1+1", Snapshot(sheet, 1, 5).Text);
        Assert.Equal(ExcelCellValueKind.Text, Snapshot(sheet, 1, 5).Kind);
    }

    [Theory]
    [InlineData(HtmlInputTrust.Untrusted)]
    [InlineData(HtmlInputTrust.Trusted)]
    public void GenericScalarsDoNotRestoreFormulaOrWorkbookStructure(HtmlInputTrust trust) {
        const string html = """
            <table><tr>
              <td data-officeimo-cell="XFD100" data-officeimo-value-kind="formula" data-officeimo-value="=1+1"><strong>Calculated elsewhere</strong></td>
              <td data-officeimo-empty="true" data-officeimo-value-kind="number" data-officeimo-value="42">42</td>
            </tr></table>
            """;
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html,
            new HtmlConversionDocumentOptions { Trust = trust }).ToExcelDocumentResult(
            new HtmlToExcelOptions {
                Mode = HtmlImportMode.Generic,
                ImportTypedCellValues = true,
                ImportFormulas = true,
                AllowUntrustedFormulas = true
            });
        using ExcelDocument workbook = result.RequireValue();
        using MemoryStream artifact = workbook.ToStream();
        using ExcelDocument reopened = ExcelDocument.Load(artifact);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);

        Assert.Equal("A1:B1", sheet.UsedRangeA1);
        Assert.Equal(ExcelCellValueKind.Text, Snapshot(sheet, 1, 1).Kind);
        Assert.Equal("Calculated elsewhere", Snapshot(sheet, 1, 1).Text);
        Assert.Equal(42D, sheet.CellAt(1, 2).GetValue<double>());
        Assert.Equal(0, result.Formulas);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticValueInvalid);
    }

    [Theory]
    [InlineData("number", "NaN")]
    [InlineData("number", "Infinity")]
    [InlineData("number", "1,25")]
    [InlineData("boolean", "perhaps")]
    [InlineData("date-time", "not-a-date")]
    [InlineData("date-time", "0001-01-02T00:00:00")]
    [InlineData("date-time", "0001-01-01T00:00:00")]
    public void InvalidScalarsKeepVisibleTextAndReportLoss(string kind, string rawValue) {
        string html = $"<table><tr><td data-officeimo-value-kind='{kind}' data-officeimo-value='{rawValue}'><strong>Unavailable</strong></td></tr></table>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, ImportTypedCellValues = true });
        using ExcelDocument workbook = result.RequireValue();

        Assert.Equal("Unavailable", Snapshot(Assert.Single(workbook.Sheets), 1, 1).Text);
        Assert.True(result.HasLoss);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticValueInvalid);
    }

    [Fact]
    public void TypedMetadataAndRichTextRespectImportLimits() {
        const string html = """
            <table>
              <tr><td data-officeimo-value-kind="number" data-officeimo-value="12345678901234567890"><strong>Short</strong></td></tr>
              <tr><td><strong>Text exceeding the field budget</strong></td></tr>
              <tr><td><strong>Past cell limit</strong></td></tr>
            </table>
            """;
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions {
                Mode = HtmlImportMode.Generic,
                ImportTypedCellValues = true,
                Limits = new HtmlImportLimits { MaxMetadataCharacters = 16, MaxTableCells = 2 }
            });
        using ExcelDocument workbook = result.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);

        Assert.Equal("Short", Snapshot(sheet, 1, 1).Text);
        Assert.False(sheet.TryGetCellValueSnapshot(2, 1, out _));
        Assert.False(sheet.TryGetCellValueSnapshot(3, 1, out _));
        Assert.Equal(1, result.Cells);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void GenericTableFormattingFollowsRowSpansAndEmptyRows() {
        const string html = """
            <style>.attention { color: #a12030; font-weight: bold; }</style>
            <table>
              <tr><th rowspan="3">Operations</th><th>Owner</th></tr>
              <tr></tr>
              <tr><td class="attention"><a href="https://example.test/review">Review required</a></td></tr>
            </table>
            """;
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.RequireValue();
        using MemoryStream artifact = workbook.ToStream();
        using ExcelDocument reopened = ExcelDocument.Load(artifact);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);

        Assert.Equal("A1:A3", Assert.Single(sheet.GetMergedRanges()).A1Range);
        Assert.Equal("Operations", Snapshot(sheet, 1, 1).Text);
        Assert.Equal("Review required", Snapshot(sheet, 3, 2).Text);
        Assert.False(sheet.TryGetCellValueSnapshot(2, 1, out _));
        Assert.False(sheet.TryGetCellValueSnapshot(3, 1, out _));
        ExcelRichTextRun run = Assert.Single(sheet.CellAt(3, 2).GetRichText());
        Assert.True(run.Bold);
        Assert.Equal("FFA12030", sheet.GetCellStyle(3, 2).FontColorArgb);
        Assert.Equal("https://example.test/review", sheet.GetHyperlinks()["B3"].Target);
    }

    [Fact]
    public void GenericCellsRetainExplicitLineBreaksWithoutRequiringOtherFormatting() {
        using ExcelDocument workbook = HtmlConversionDocument.Parse("<table><tr><td>First line<br>Second line</td></tr></table>")
            .ToExcelDocument(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        Assert.Equal("First line\nSecond line", Snapshot(Assert.Single(workbook.Sheets), 1, 1).Text);
    }

    private static ExcelCellValueSnapshot Snapshot(ExcelSheet sheet, int row, int column) {
        Assert.True(sheet.TryGetCellValueSnapshot(row, column, out ExcelCellValueSnapshot? value));
        return value!;
    }
}
