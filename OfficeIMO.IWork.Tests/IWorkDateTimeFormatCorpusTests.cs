using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkDateTimeFormatCorpusTests {
    [Fact]
    public void Unchanged_native_date_cells_match_independent_patterns_caches_and_display() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "numbers-parser", "date-formats.json")));
        foreach (var package in manifest.RootElement.GetProperty("packages").EnumerateArray()) {
            string path = Path.Combine(root, package.GetProperty("source").GetString()!);
            Assert.Equal(package.GetProperty("sourceSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
            var projection = IWorkSourceDocument.Open(path, new IWorkReadOptions { PreserveSourceRecords = false }).ReadNumbers();
            foreach (var expected in package.GetProperty("cells").EnumerateArray()) {
                var table = projection.Sheets.Single(s => s.Name == expected.GetProperty("sheet").GetString())
                    .Tables.Single(t => t.Name == expected.GetProperty("table").GetString());
                var cell = table.GetCell(expected.GetProperty("row").GetInt32(), expected.GetProperty("column").GetInt32())!;
                Assert.Equal(DateTime.Parse(expected.GetProperty("value").GetString()!, CultureInfo.InvariantCulture),
                    Assert.IsType<DateTime>(cell.Value));
                Assert.Equal(IWorkCellKind.DateTime, cell.ValueKind);
                Assert.False(cell.HasDecodeError);
                Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures & IWorkCellUnsupportedFeatures.DateFormat);
                Assert.Equal(expected.GetProperty("settings").GetProperty("date_time_format").GetString(),
                    cell.NumberFormat!.DateTimeFormat!.SourcePattern);
                Assert.True(cell.TryGetFormattedNumber(out string display, out _));
                Assert.Equal(expected.GetProperty("independentDisplayText").GetString(), display);
            }
        }
    }

    [Fact]
    public void Saved_xlsx_preserves_original_date_formula_caches_and_reports_locale_limits() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "numbers-parser", "date-formats.json")));
        var package = manifest.RootElement.GetProperty("packages").EnumerateArray()
            .Single(p => p.GetProperty("source").GetString() == "numbers-parser/test-10-formulas.numbers");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(Path.Combine(root, package.GetProperty("source").GetString()!));
        Assert.False(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_DATE_DISPLAY_APPROXIMATED");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        foreach (var expected in package.GetProperty("cells").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
            var source = result.Projection.Sheets[0].Tables[0].GetCell(row, column)!;
            var actual = reopened.Sheets[0].CellAt(row, column);
            Assert.Equal(Assert.IsType<DateTime>(source.Value).ToOADate(), actual.GetValue<double>());
            Assert.Equal(source.NumberFormat!.ToSpreadsheetFormatCode(), actual.GetStyle().NumberFormatCode);
            Assert.True(source.FormulaIsComplete);
            Assert.Equal(source.Formula!.TrimStart('='), actual.GetValue().Formula);
            Assert.Equal(expected.GetProperty("independentDisplayText").GetString(),
                Assert.Single(reopened.Sheets[0].Range(actual.Address).CreateVisualSnapshot().Cells).Text);
        }
        var export = manifest.RootElement.GetProperty("nativeExport");
        string nativePath = Path.Combine(root, export.GetProperty("path").GetString()!);
        Assert.Equal(export.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(nativePath))).ToLowerInvariant());
        using var native = ExcelDocument.Load(nativePath);
        foreach (var expected in export.GetProperty("cells").EnumerateArray()) {
            var actual = native.Sheets[0].Range(expected.GetProperty("cell").GetString()!).FirstCell;
            Assert.Equal(expected.GetProperty("valueDays").GetDouble(), actual.GetValue<double>());
            Assert.Equal(expected.GetProperty("formatCode").GetString(), actual.GetStyle().NumberFormatCode);
        }
    }
}
