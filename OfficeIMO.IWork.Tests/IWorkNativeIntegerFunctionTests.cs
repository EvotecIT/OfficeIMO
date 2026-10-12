using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(58, new double[] { 999999999999999, 1 }, true, 1d)]
    [InlineData(58, new double[] { 1000000000000000, 1 }, false, 42d)]
    [InlineData(75, new double[] { 999999999999999, 1 }, true, 999999999999999d)]
    [InlineData(75, new double[] { 1000000000000000, 1 }, false, 42d)]
    [InlineData(75, new double[] { 100000000, 10000000 }, false, 42d)]
    [InlineData(75, new double[] { 99999999, 10000000 }, true, 999999990000000d)]
    [InlineData(75, new double[] { 100000000, 10000000, 0 }, true, 0d)]
    [InlineData(75, new double[] { 0, 100000000, 10000000 }, true, 0d)]
    public void Native_integer_function_literal_bounds_control_editability_after_saved_conversion(
        int index, double[] operands, bool qualified, double expected) {
        var nodes = operands.Select(number => BytesField(1,
            Message(VarintField(1, 17), DoubleField(4, number)))).ToList();
        nodes.Add(BytesField(1, Message(VarintField(1, 16), VarintField(2, (ulong)index),
            VarintField(3, (ulong)operands.Length))));
        using var package = CreateNumbersPackage(new[] { new TableSpec("Integers", 1, 1, 42d,
            hasFormula: true, formulaPayload: BytesField(1, Message(nodes.ToArray()))) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.Equal(qualified, result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var sheet = Assert.Single(reopened.Sheets);
        Assert.Equal(42d, sheet.CellAt(1, 1).GetValue<double>());
        Assert.Equal(qualified, sheet.GetFormulaText(1, 1) != null);
        Assert.Equal(qualified ? 1 : 0, reopened.Calculate());
        Assert.Equal(expected, sheet.CellAt(1, 1).GetValue<double>());
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Fact]
    public void Native_integer_functions_reconstruct_qualified_literals_and_numeric_ranges_and_retain_incompatible_caches() {
        using var manifest = ReadNativeIntegerFunctionManifest();
        JsonElement evidence = manifest.RootElement;
        var qualified = new HashSet<string>(new[] { "G2", "H2", "G3", "H3", "G4", "H4", "G5", "H5", "G6", "H6",
            "G12", "H12", "G15", "G16", "H16", "G17", "H17", "G18", "H18", "B13", "B14",
            "E2", "F2", "E4", "F4", "E5", "F5", "E6", "F6", "E17", "F17", "E18", "F18" });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(CorpusFixture("native-exports/" + evidence.GetProperty("source").GetString()),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        IWorkTable table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet destination = Assert.Single(reopened.Sheets);
        foreach (JsonElement expected in evidence.GetProperty("cells").EnumerateArray().Where(c => c.GetProperty("formula").ValueKind == JsonValueKind.String)) {
            int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
            string address = expected.GetProperty("address").GetString()!;
            IWorkTableCell source = table.GetCell(row, column)!;
            Assert.True(source.CachedValueIsComplete);
            Assert.Equal(qualified.Contains(address), source.FormulaIsComplete);
            if (qualified.Contains(address)) Assert.Equal(expected.GetProperty("formula").GetString(), destination.GetFormulaText(row, column));
            else Assert.Null(destination.GetFormulaText(row, column));
            JsonElement native = expected.GetProperty("nativeValue");
            if (native.ValueKind == JsonValueKind.Number) {
                Assert.Equal(native.GetDouble(), Assert.IsType<double>(source.Value));
                Assert.Equal(native.GetDouble(), destination.CellAt(row, column).GetValue<double>());
            } else {
                Assert.Equal(IWorkCellKind.Error, source.ValueKind);
                Assert.Equal("#ERROR", source.Value);
                Assert.Equal("#ERROR", destination.CellAt(row, column).GetValue<string>());
            }
        }
        Assert.Equal(qualified.Count, reopened.Calculate());
        foreach (JsonElement expected in evidence.GetProperty("cells").EnumerateArray()
            .Where(c => qualified.Contains(c.GetProperty("address").GetString()!) && c.GetProperty("column").GetInt32() >= 5)) {
            Assert.Equal(expected.GetProperty("nativeValue").GetDouble(), destination.CellAt(
                expected.GetProperty("row").GetInt32(), expected.GetProperty("column").GetInt32()).GetValue<double>());
        }
        destination.CellValue(2, 2, 30d);
        reopened.Calculate(); Assert.Equal(6d, destination.CellAt(2, 5).GetValue<double>());
        Assert.Equal(180d, destination.CellAt(2, 6).GetValue<double>());
        using var edited = new MemoryStream(); reopened.Save(edited); edited.Position = 0;
        using var final = ExcelDocument.Load(edited);
        Assert.Equal(6d, final.Sheets[0].CellAt(2, 5).GetValue<double>());
        Assert.Equal(180d, final.Sheets[0].CellAt(2, 6).GetValue<double>());
        Assert.Empty(final.ValidateOpenXml());
    }

    [Fact]
    public void Shared_integer_evaluation_matches_fresh_Excel_including_native_Numbers_coercion_differences() {
        using var manifest = ReadNativeIntegerFunctionManifest();
        JsonElement evidence = manifest.RootElement;
        using var document = ExcelDocument.Load(CorpusFixture("native-exports/" + evidence.GetProperty("nativeXlsx").GetString()));
        ExcelSheet sheet = Assert.Single(document.Sheets);
        Assert.Equal(68, document.Calculate());
        foreach (JsonElement cell in evidence.GetProperty("cells").EnumerateArray().Where(c => c.GetProperty("formula").ValueKind == JsonValueKind.String)) {
            int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
            // 2^53 is outside the evaluator's documented operand domain. Do not
            // turn this one Excel version's acceptance into a wider guarantee.
            if (row == 14 && column is 5 or 6) continue;
            JsonElement expected = cell.GetProperty("excelRecalculatedXlsx");
            Assert.True(sheet.TryGetCachedFormulaValue(row, column, out string? actual));
            if (expected.GetProperty("kind").GetString() == "e") Assert.Equal(expected.GetProperty("value").GetString(), actual);
            else Assert.Equal(double.Parse(expected.GetProperty("value").GetString()!, CultureInfo.InvariantCulture),
                double.Parse(actual!, CultureInfo.InvariantCulture));
        }
        Assert.Equal(360d, sheet.CellAt(10, 6).GetValue<double>()); // Native Numbers cached zero for this numeric-text range.
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(360d, reopened.Sheets[0].CellAt(10, 6).GetValue<double>());
        Assert.Empty(reopened.ValidateOpenXml());
    }

    private static JsonDocument ReadNativeIntegerFunctionManifest() {
        var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("native-exports/numbers-integer-functions-v14.5-decode.json")));
        foreach (JsonElement input in manifest.RootElement.GetProperty("inputs").EnumerateArray()) {
            string path = CorpusFixture("native-exports/" + input.GetProperty("path").GetString());
            Assert.Equal(input.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        }
        return manifest;
    }
}
