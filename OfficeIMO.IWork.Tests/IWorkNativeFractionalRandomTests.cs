using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_fractional_random_bounds_retain_caches_and_evaluate_the_qualified_integer_intersection() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("native-exports/numbers-fractional-bounds-v14.5.json")));
        var fixture = Assert.Single(manifest.RootElement.GetProperty("fixtures").EnumerateArray());
        foreach (var artifact in fixture.GetProperty("artifacts").EnumerateArray())
            Assert.Equal(artifact.GetProperty("sha256").GetString(), HashFile(CorpusFixture("native-exports/" + artifact.GetProperty("path").GetString())));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(CorpusFixture("native-exports/" + fixture.GetProperty("source").GetString()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        var sourceSheet = Assert.Single(result.Projection.Sheets);
        Assert.Single(sourceSheet.Drawables); // No unrelated text boxes in the native fixture.
        var table = Assert.Single(sourceSheet.Tables);
        Assert.Equal(17, table.RowCount); Assert.Equal(5, table.ColumnCount);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var sheet = Assert.Single(reopened.Sheets);
        foreach (var expected in fixture.GetProperty("cells").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            var cell = table.GetCell(row, 2)!;
            string formula = expected.GetProperty("formula").GetString()!;
            Assert.True(cell.FormulaIsComplete);
            Assert.Equal("=" + formula, cell.Formula);
            Assert.Equal(formula, sheet.GetFormulaText(row, 2));
            var cache = expected.GetProperty("nativeNumericCache");
            if (cache.ValueKind == JsonValueKind.Number) Assert.Equal(cache.GetDouble(), sheet.CellAt(row, 2).GetValue<double>());
            else Assert.Equal("#ERROR", sheet.CellAt(row, 2).GetValue<string>());
        }
        Assert.Equal(9, reopened.Calculate()); // Seven nonempty intersections, two reversed intervals.
        foreach (var expected in fixture.GetProperty("cells").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            string profile = expected.GetProperty("localEvaluation").GetString()!;
            if (profile == "reversed-bounds") {
                Assert.Equal("#NUM!", expected.GetProperty("nativeExcelCache").GetString());
                Assert.True(sheet.TryGetCachedFormulaValue(row, 2, out string? error)); Assert.Equal("#NUM!", error);
            } else {
                double native = expected.GetProperty("nativeNumericCache").GetDouble();
                Assert.Equal(native, sheet.CellAt(row, 2).GetValue<double>());
                if (profile == "integer-interval") Assert.Equal(native, expected.GetProperty("nativeExcelCache").GetDouble());
            }
        }
        // An edit within the supported operand profile produces a fresh result.
        sheet.CellValue(10, 3, -2.9d); sheet.CellValue(10, 4, -1.1d);
        sheet.CellFormula(10, 2, "RANDBETWEEN(C10,D10)"); reopened.Calculate();
        Assert.Equal(-2d, sheet.CellAt(10, 2).GetValue<double>());
        using var calculated = new MemoryStream(); reopened.Save(calculated); calculated.Position = 0;
        using var final = ExcelDocument.Load(calculated);
        Assert.Equal(-2d, final.Sheets[0].CellAt(10, 2).GetValue<double>());
        Assert.Empty(final.ValidateOpenXml());
    }
}
