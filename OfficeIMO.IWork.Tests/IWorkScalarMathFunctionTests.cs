using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(113, "PRODUCT", 5, 120d)]
    [InlineData(117, "RADIANS", 1, Math.PI / 180d)]
    [InlineData(44, "DEGREES", 1, 180d / Math.PI)]
    [InlineData(51, "FACT", 1, 1d)]
    [InlineData(48, "EVEN", 1, 2d)]
    [InlineData(100, "ODD", 1, 1d)]
    [InlineData(147, "SUMSQ", 2, 5d)]
    [InlineData(17, "CEILING", 2, 2d)]
    [InlineData(55, "FLOOR", 2, 0d)]
    [InlineData(65, "INT", 1, 1d)]
    [InlineData(92, "MOD", 2, 1d)]
    [InlineData(139, "SQRT", 1, 1d)]
    [InlineData(127, "ROUNDDOWN", 2, 1d)]
    [InlineData(128, "ROUNDUP", 2, 1d)]
    [InlineData(133, "SIGN", 1, 1d)]
    [InlineData(157, "TRUNC", 1, 1d)]
    [InlineData(157, "TRUNC", 2, 1d)]
    [InlineData(50, "EXP", 1, Math.E)]
    [InlineData(78, "LN", 1, 0d)]
    [InlineData(79, "LOG", 1, 0d)]
    [InlineData(79, "LOG", 2, 0d)]
    [InlineData(80, "LOG10", 1, 0d)]
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
    [InlineData(17, new double[] { 2.5, -2 }, "CEILING(2.5,-2)", "#NUM!")]
    [InlineData(55, new double[] { 1, 0 }, "FLOOR(1,0)", "#DIV/0!")]
    [InlineData(51, new double[] { -1 }, "FACT(-1)", "#NUM!")]
    [InlineData(51, new double[] { 171 }, "FACT(171)", "#NUM!")]
    [InlineData(139, new double[] { -1 }, "SQRT(-1)", "#NUM!")]
    [InlineData(92, new double[] { 1, 0 }, "MOD(1,0)", "#DIV/0!")]
    [InlineData(50, new double[] { 1000 }, "EXP(1000)", "#NUM!")]
    [InlineData(78, new double[] { 0 }, "LN(0)", "#NUM!")]
    [InlineData(79, new double[] { 8, 1 }, "LOG(8,1)", "#DIV/0!")]
    [InlineData(80, new double[] { -1 }, "LOG10(-1)", "#NUM!")]
    public void Converted_scalar_math_domain_errors_replace_numeric_caches_in_saved_xlsx(int index, double[] operands, string expression, string error) {
        var nodes = operands.Select(operand => BytesField(1, Message(VarintField(1, 17), DoubleField(4, operand)))).ToList();
        nodes.Add(BytesField(1, Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, (ulong)operands.Length))));
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
    [InlineData(117, 1)]
    [InlineData(44, 1)]
    [InlineData(51, 1)]
    [InlineData(48, 1)]
    [InlineData(100, 1)]
    [InlineData(17, 2)]
    [InlineData(55, 2)]
    [InlineData(65, 1)]
    [InlineData(92, 2)]
    [InlineData(139, 1)]
    [InlineData(127, 2)]
    [InlineData(128, 2)]
    [InlineData(133, 1)]
    [InlineData(50, 1)]
    [InlineData(78, 1)]
    [InlineData(80, 1)]
    public void Scalar_math_native_functions_require_exact_argument_counts(int index, int count) {
        Assert.True(RenderFunction(index, count).IsComplete);
        Assert.False(RenderFunction(index, count - 1).IsComplete);
        Assert.False(RenderFunction(index, count + 1).IsComplete);
    }
    [Theory]
    [InlineData(127, 31415.92654d, -2d, "ROUNDDOWN(31415.92654,-2)", 31400d)]
    [InlineData(128, -3.14159d, 1d, "ROUNDUP(-3.14159,1)", -3.2d)]
    [InlineData(157, -1234.5d, -2d, "TRUNC(-1234.5,-2)", -1200d)]
    [InlineData(133, -0.00001d, null, "SIGN(-1E-05)", -1d)]
    [InlineData(51, 5.9d, null, "FACT(5.9)", 120d)]
    [InlineData(48, -1.5d, null, "EVEN(-1.5)", -2d)]
    [InlineData(100, -2d, null, "ODD(-2)", -3d)]
    public void Directional_math_native_functions_recalculate_signed_operands_after_saved_conversion(int index, double operand, double? digits, string expression, double expected) {
        var nodes = new List<byte[]> { BytesField(1, Message(VarintField(1, 17), DoubleField(4, operand))) };
        if (digits.HasValue) nodes.Add(BytesField(1, Message(VarintField(1, 17), DoubleField(4, digits.Value))));
        nodes.Add(BytesField(1, Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, digits.HasValue ? 2ul : 1ul))));
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, 42d,
            hasFormula: true, formulaPayload: BytesField(1, Message(nodes.ToArray()))) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.Equal("=" + expression, result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.Formula);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(expression, reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Fact]
    public void Trunc_native_function_accepts_only_its_required_and_optional_argument_counts() {
        Assert.False(RenderFunction(157, 0).IsComplete);
        Assert.True(RenderFunction(157, 1).IsComplete);
        Assert.True(RenderFunction(157, 2).IsComplete);
        Assert.False(RenderFunction(157, 3).IsComplete);
    }

    [Fact]
    public void Log_native_function_accepts_only_its_required_and_optional_argument_counts() {
        Assert.False(RenderFunction(79, 0).IsComplete);
        Assert.True(RenderFunction(79, 1).IsComplete);
        Assert.True(RenderFunction(79, 2).IsComplete);
        Assert.False(RenderFunction(79, 3).IsComplete);
    }
}
