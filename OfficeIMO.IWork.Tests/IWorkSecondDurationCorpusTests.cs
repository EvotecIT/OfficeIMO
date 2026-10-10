using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkSecondDurationCorpusTests {
    [Fact]
    public void Native_fixed_second_ranges_match_Apple_codes_caches_and_recalculate_saved_formulas() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "numbers-duration-seconds-v14.5.json")));
        foreach (var artifact in manifest.RootElement.GetProperty("artifacts").EnumerateArray()) {
            Assert.Equal(artifact.GetProperty("sha256").GetString(), Hash(Path.Combine(root, artifact.GetProperty("path").GetString()!)));
        }
        using var decoded = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, manifest.RootElement.GetProperty("decode").GetString()!)));
        string sourcePath = Path.Combine(root, decoded.RootElement.GetProperty("source").GetString()!);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(sourcePath);
        result.Report.RequireCompleteEditableReconstruction();
        Assert.False(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED");
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code is "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED" or "IWORK_TABLE_CELL_FEATURES_UNASSESSED");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        using var native = ExcelDocument.Load(Path.Combine(root, decoded.RootElement.GetProperty("nativeExport").GetString()!));
        Assert.Equal(3, result.Projection.Sheets.Count);
        int cells = 0, formulas = 0, integralDisplays = 0;
        foreach (var profile in decoded.RootElement.GetProperty("profiles").EnumerateArray()) {
            string sheetName = profile.GetProperty("sheet").GetString()!;
            var table = Assert.Single(result.Projection.Sheets.Single(s => s.Name == sheetName).Tables);
            Assert.Equal(profile.GetProperty("rows").GetInt32(), table.RowCount);
            Assert.Equal(profile.GetProperty("columns").GetInt32(), table.ColumnCount);
            var target = reopened.Sheets.Single(s => s.Name == sheetName);
            var reference = native.Sheets.Single(s => s.Name == sheetName);
            var appearance = manifest.RootElement.GetProperty("profiles").EnumerateArray().Single(p => p.GetProperty("sheet").GetString() == sheetName);
            foreach (var expected in profile.GetProperty("cells").EnumerateArray()) {
                int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
                string address = expected.GetProperty("address").GetString()!;
                var cell = table.GetCell(row, column)!;
                var format = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat);
                Assert.Equal(IWorkCellKind.Duration, cell.ValueKind);
                Assert.Equal(expected.GetProperty("seconds").GetDouble(), Assert.IsType<double>(cell.Value));
                Assert.True(cell.CachedValueIsComplete);
                Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
                Assert.Equal((IWorkDurationUnit)expected.GetProperty("settings").GetProperty("duration_unit_largest").GetInt32(), format.DurationFormat!.LargestUnit);
                Assert.Equal(IWorkDurationUnit.Second, format.DurationFormat.SmallestUnit);
                Assert.Equal(expected.GetProperty("nativeFormatCode").GetString(), target.CellAt(row, column).GetStyle().NumberFormatCode);
                Assert.Equal(reference.CellAt(row, column).GetStyle().NumberFormatCode, target.CellAt(row, column).GetStyle().NumberFormatCode);
                Assert.InRange(Math.Abs(expected.GetProperty("nativeValueDays").GetDouble() - target.CellAt(row, column).GetValue<double>()), 0, 1e-12);
                Assert.InRange(Math.Abs(reference.CellAt(row, column).GetValue<double>() - target.CellAt(row, column).GetValue<double>()), 0, 1e-12);
                if (expected.GetProperty("isFormula").GetBoolean()) {
                    Assert.Equal(IWorkCellKind.Formula, cell.Kind);
                    Assert.True(cell.SourceFormulaIsDeclared);
                    Assert.True(cell.FormulaIsComplete);
                    Assert.Equal(expected.GetProperty("formula").GetString(), cell.Formula!.TrimStart('='));
                    Assert.Equal(expected.GetProperty("nativeFormula").GetString(), target.GetFormulaText(row, column));
                    formulas++;
                }
                // Fractional-second appearance is explicitly approximated. The
                // native PDF still independently qualifies all integral rows.
                if (row <= 6) {
                    string display = appearance.GetProperty("cells").EnumerateArray().Single(c => c.GetProperty("address").GetString() == address)
                        .GetProperty("numbersPdfDisplay").GetString()!;
                    Assert.True(cell.TryGetFormattedNumber(out string text, out _));
                    Assert.Equal(display, text);
                    Assert.Equal(display, Assert.Single(target.Range(address).CreateVisualSnapshot().Cells).Text);
                    integralDisplays++;
                }
                cells++;
            }
            target.CellValue(5, 3, 1.75d); // Edit the operand in day serials, retaining the native duration format.
        }
        Assert.Equal(42, cells);
        Assert.Equal(21, formulas);
        Assert.Equal(30, integralDisplays);
        Assert.True(reopened.Calculate() > 0);
        using var recalculated = new MemoryStream(); reopened.Save(recalculated); recalculated.Position = 0;
        using var final = ExcelDocument.Load(recalculated);
        foreach (var sheet in final.Sheets) {
            Assert.Equal("C5+C5", sheet.GetFormulaText(5, 4));
            Assert.Equal(3.5d, sheet.CellAt(5, 4).GetValue<double>());
        }
        Assert.Empty(final.ValidateOpenXml());
        Assert.Equal(decoded.RootElement.GetProperty("sourceSha256").GetString(), Hash(sourcePath));
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
