using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(23, "COLUMNS")]
    [InlineData(130, "ROWS")]
    public void Reference_shape_literal_arrays_preserve_editability_and_cache_when_local_evaluation_is_unsupported(int index, string name) {
        byte[] formula = SingleCellFormula(
            Message(VarintField(1, 17), DoubleField(4, 1)),
            Message(VarintField(1, 17), DoubleField(4, 2)),
            Message(VarintField(1, 17), DoubleField(4, 3)),
            Message(VarintField(1, 24), VarintField(11, 3), VarintField(12, 1)),
            Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, 1)));
        using var package = CreateNumbersPackage(new[] { new TableSpec("Shape", 1, 1, 42d,
            hasFormula: true, formulaPayload: formula) });
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(name + "({1,2,3})", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(0, reopened.Calculate());
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.False(Assert.Single(reopened.InspectFormulas().Formulas).IsSupportedByOfficeIMO);
    }

    [Theory]
    [InlineData(129, "ROW", false, 1d)]
    [InlineData(130, "ROWS", false, 1d)]
    [InlineData(23, "COLUMNS", false, 1d)]
    [InlineData(130, "ROWS", true, 3d)]
    [InlineData(23, "COLUMNS", true, 2d)]
    public void Reference_shape_functions_keep_bound_coordinates_and_recalculate_after_save(
        int index, string name, bool range, double expected) {
        var nodes = new List<byte[]> { SingleCellReference(true, true) };
        if (range) {
            nodes.Add(Message(VarintField(1, 36),
                BytesField(26, Message(VarintField(1, 2), VarintField(2, 1))),
                BytesField(27, Message(VarintField(1, 4), VarintField(2, 1))), SingleCellIdentity()));
            nodes.Add(VarintField(1, 29));
        }
        nodes.Add(Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, 1)));
        using var package = CrossBindingPackage(formulaPayload: SingleCellFormula(nodes.ToArray()));
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.False(converted.IsVisualFallback);
        Assert.True(converted.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        string target = converted.WorksheetMappings[1].DestinationName.Replace("'", "''");
        Assert.Equal(name + "('" + target + "'!$A$1" + (range ? ":$B$3" : "") + ")",
            reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        reopened.Sheets[1].CellValue(1, 1, 999d);
        reopened.Calculate();
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Theory]
    [InlineData(129, "ROW", 0, 1d)]
    [InlineData(22, "COLUMN", 0, 1d)]
    [InlineData(130, "ROWS", 1, 1d)]
    [InlineData(23, "COLUMNS", 1, 1d)]
    public void Reference_shape_native_functions_recalculate_current_cell_or_scalar_operands(
        int index, string name, int arguments, double expected) {
        using var package = CreateNumbersPackage(new[] { new TableSpec("Shape", 1, 1, 42d,
            hasFormula: true, formulaPayload: FunctionPayload(index, arguments)) });
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(name + "(" + (arguments == 0 ? "" : "1") + ")", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.False(RenderFunction(index, 2).IsComplete);
        if (arguments == 1) Assert.False(RenderFunction(index, 0).IsComplete);
    }
}
