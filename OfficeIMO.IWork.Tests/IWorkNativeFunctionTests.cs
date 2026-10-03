using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_Numbers_function_expressions_and_caches_match_Apple_exports_after_save_and_recalculation() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("native-exports/numbers-functions-v14.5.json")));
        foreach (var fixture in manifest.RootElement.GetProperty("fixtures").EnumerateArray()) {
            foreach (var artifact in fixture.GetProperty("artifacts").EnumerateArray())
                Assert.Equal(artifact.GetProperty("sha256").GetString(), HashFile(CorpusFixture("native-exports/" + artifact.GetProperty("path").GetString())));
            using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(CorpusFixture("native-exports/" + fixture.GetProperty("source").GetString()),
                conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
            Assert.False(result.IsVisualFallback);
            Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
            var table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
            Assert.Equal(fixture.GetProperty("rows").GetInt32(), table.RowCount);
            Assert.Equal(fixture.GetProperty("columns").GetInt32(), table.ColumnCount);
            using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            var destination = Assert.Single(reopened.Sheets);
            foreach (var expected in fixture.GetProperty("cells").EnumerateArray()) {
                int row = expected.GetProperty("row").GetInt32();
                IWorkTableCell cell = table.GetCell(row, 5)!;
                string formula = expected.GetProperty("formula").GetString()!;
                Assert.True(cell.FormulaIsComplete); Assert.True(cell.CachedValueIsComplete);
                Assert.Equal("=" + formula, cell.Formula);
                Assert.Equal(formula, destination.GetFormulaText(row, 5));
                var cache = expected.GetProperty("nativeNumericCache");
                if (cache.ValueKind == JsonValueKind.Number) {
                    Assert.Equal(cache.GetDouble(), cell.Value);
                    Assert.Equal(cache.GetDouble(), destination.CellAt(row, 5).GetValue<double>());
                } else {
                    // Native Numbers emits no XLSX error code for these cells. Retain
                    // the diagnosed source display cache rather than inventing equivalence.
                    Assert.Equal("#ERROR", cell.Value);
                    Assert.Equal("#ERROR", destination.CellAt(row, 5).GetValue<string>());
                    Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_ERROR_VALUE_APPROXIMATED");
                }
            }
            int expectedCount = fixture.GetProperty("cells").GetArrayLength() == 9 ? 9 : 15;
            Assert.Equal(expectedCount, reopened.Calculate());
            foreach (var expected in fixture.GetProperty("cells").EnumerateArray()) {
                int row = expected.GetProperty("row").GetInt32();
                string formula = expected.GetProperty("formula").GetString()!;
                var cache = expected.GetProperty("nativeNumericCache");
                if (cache.ValueKind != JsonValueKind.Number || formula == "RANDBETWEEN(1,10)") continue;
                Assert.Equal(cache.GetDouble(), destination.CellAt(row, 5).GetValue<double>(), precision: 12);
            }
            Assert.InRange(destination.CellAt(8, 5).GetValue<double>(), 1d, 10d);
            destination.CellValue(4, 1, 20d); reopened.Calculate();
            Assert.Equal(20d, destination.CellAt(5, 5).GetValue<double>());
            Assert.Equal(24d, destination.CellAt(6, 5).GetValue<double>());
            Assert.Equal(0d, destination.CellAt(2, 5).GetValue<double>());
            using var recalculated = new MemoryStream(); reopened.Save(recalculated); recalculated.Position = 0;
            using var final = ExcelDocument.Load(recalculated);
            Assert.Equal(24d, final.Sheets[0].CellAt(6, 5).GetValue<double>());
            Assert.True(final.Sheets[0].TryGetCachedFormulaValue(9, 5, out string? unequal));
            Assert.Equal("#N/A", unequal); // Excel's evaluation contract, not a native error-cache mapping.
        }
    }

    [Theory]
    [InlineData(101, 3, 5)]
    [InlineData(112, 3, 4)]
    [InlineData(119, 2, 2)]
    public void Qualified_native_functions_enforce_their_own_argument_counts(int index, int minimum, int maximum) {
        for (int count = minimum; count <= maximum; count++) Assert.True(RenderFunction(index, count).IsComplete);
        Assert.False(RenderFunction(index, minimum - 1).IsComplete);
        Assert.False(RenderFunction(index, maximum + 1).IsComplete);
    }
}
