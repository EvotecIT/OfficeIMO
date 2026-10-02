using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(65, "INT", 1, 1d)]
    [InlineData(92, "MOD", 2, 1d)]
    [InlineData(139, "SQRT", 1, 1d)]
    public void Scalar_math_native_functions_save_reopen_and_recalculate_numeric_caches(int index, string name, int count, double expected) {
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, 42d,
            hasFormula: true, formulaPayload: FunctionPayload(index, count)) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        var cell = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        string expression = name + "(" + string.Join(",", Enumerable.Range(1, count)) + ")";
        Assert.True(cell.FormulaIsComplete);
        Assert.Equal("=" + expression, cell.Formula);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var sheet = reopened.Sheets[0];
        Assert.Equal(expression, sheet.GetFormulaText(1, 1));
        Assert.Equal(42d, sheet.CellAt(1, 1).GetValue<double>());
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(expected, sheet.CellAt(1, 1).GetValue<double>());
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Theory]
    [InlineData(139, "SQRT(-1)", "#NUM!")]
    [InlineData(92, "MOD(1,0)", "#DIV/0!")]
    public void Converted_scalar_math_domain_errors_replace_numeric_caches_in_saved_xlsx(int index, string expression, string error) {
        var nodes = new List<byte[]> {
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, index == 139 ? -1d : 1d)))
        };
        if (index == 92) nodes.Add(BytesField(1, Message(VarintField(1, 17), DoubleField(4, 0d))));
        nodes.Add(BytesField(1, Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, index == 92 ? 2ul : 1ul))));
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, 42d,
            hasFormula: true, formulaPayload: BytesField(1, Message(nodes.ToArray()))) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.Equal("=" + expression, result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.Formula);
        Assert.Equal(42d, result.Value.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(1, result.Value.Calculate());
        IWorkFormulaCacheAssertions.HasType(result.Value, 0, "A1", "e");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(expression, reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(error, Assert.Single(reopened.InspectFormulas().Formulas).CachedValue);
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Theory]
    [InlineData(65, 1)]
    [InlineData(92, 2)]
    [InlineData(139, 1)]
    public void Scalar_math_native_functions_require_exact_argument_counts(int index, int count) {
        Assert.True(RenderFunction(index, count).IsComplete);
        Assert.False(RenderFunction(index, count - 1).IsComplete);
        Assert.False(RenderFunction(index, count + 1).IsComplete);
    }
}
