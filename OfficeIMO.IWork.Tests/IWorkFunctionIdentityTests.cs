using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Function_identifiers_match_the_independent_provider_for_the_supported_subset() {
        using var manifest = ReadFunctionIdentityManifest();
        foreach (JsonElement identity in manifest.RootElement.GetProperty("identities").EnumerateArray()) {
            int arguments = identity.GetProperty("sampleArgumentCount").GetInt32();
            IWorkFormulaResult result = RenderFunction(identity.GetProperty("index").GetInt32(), arguments);
            Assert.True(result.IsComplete);
            string expected = "=" + identity.GetProperty("name").GetString() + "("
                + string.Join(",", Enumerable.Range(1, arguments)) + ")";
            Assert.Equal(expected, result.Text);
        }
    }

    [Theory]
    [InlineData(89, 1)] // Native MINA; previously exported as MINUTE.
    [InlineData(101, 3)] // Native OFFSET; previously exported as OR.
    [InlineData(112, 2)] // PROB cannot use ROUND's signature.
    [InlineData(119, 1)] // RANDBETWEEN cannot use SECOND's signature.
    [InlineData(169, 2)] // No qualified native SUMIF identity here.
    public void Unqualified_function_identifiers_preserve_caches_without_an_editable_formula(int index, int arguments) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Functions", 1, 1, 42d, hasFormula: true,
                formulaPayload: FunctionPayload(index, arguments))
        });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell source = Assert.Single(Assert.Single(Assert.Single(result.Projection.Sheets).Tables).Cells);
        Assert.False(source.FormulaIsComplete);
        Assert.True(source.CachedValueIsComplete);
        Assert.Equal(42d, source.Value);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        Assert.Null(reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Native_nested_or_and_power_formulas_keep_their_identities_and_saved_typed_caches() {
        using var manifest = ReadFunctionIdentityManifest();
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser", "cross-table-formulas.numbers");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        foreach (JsonElement expected in manifest.RootElement.GetProperty("nativeCases").EnumerateArray()) {
            string sheetName = expected.GetProperty("sourceSheet").GetString()!;
            string tableName = expected.GetProperty("sourceTable").GetString()!;
            int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
            IWorkTable table = Assert.Single(Assert.Single(result.Projection.Sheets, s => s.Name == sheetName).Tables, t => t.Name == tableName);
            IWorkTableCell cell = table.GetCell(row, column)!;
            string expression = expected.GetProperty("formula").GetString()!;
            Assert.True(cell.FormulaIsComplete);
            Assert.Equal("=" + expression, cell.Formula);
            NumbersWorksheetMapping mapping = Assert.Single(result.WorksheetMappings,
                m => m.SourceSheetName == sheetName && m.SourceTableName == tableName);
            ExcelSheet destination = Assert.Single(reopened.Sheets, s => s.Name == mapping.DestinationName);
            Assert.Equal(expression, destination.GetFormulaText(row, column));
            JsonElement cache = expected.GetProperty("cachedValue");
            if (cache.ValueKind == JsonValueKind.String) {
                Assert.Equal(cache.GetString(), cell.Value);
                Assert.Equal(cache.GetString(), destination.CellAt(row, column).GetValue<string>());
            } else {
                Assert.Equal(cache.GetDouble(), cell.Value);
                Assert.Equal(cache.GetDouble(), destination.CellAt(row, column).GetValue<double>());
            }
        }
    }

    private static JsonDocument ReadFunctionIdentityManifest() => JsonDocument.Parse(File.ReadAllText(
        Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser", "function-identities.json")));

    private static byte[] FunctionPayload(int index, int argumentCount) {
        var nodes = new List<byte[]>(argumentCount + 1);
        for (int argument = 0; argument < argumentCount; argument++) {
            nodes.Add(BytesField(1, Message(VarintField(1, 17), DoubleField(4, argument + 1d))));
        }
        nodes.Add(BytesField(1, Message(VarintField(1, 16), VarintField(2, checked((ulong)index)),
            VarintField(3, checked((ulong)argumentCount)))));
        return BytesField(1, Message(nodes.ToArray()));
    }
}
