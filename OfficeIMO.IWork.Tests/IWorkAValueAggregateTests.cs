using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(16, "AVERAGEA", 1.5d)]
    [InlineData(85, "MAXA", 2d)]
    [InlineData(89, "MINA", 1d)]
    public void A_value_aggregate_native_identity_saves_an_editable_formula_and_numeric_cache(int index, string name, double expected) {
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, expected,
            hasFormula: true, formulaPayload: FunctionPayload(index, 2)) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        var source = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.True(source.FormulaIsComplete);
        Assert.Equal("=" + name + "(1,2)", source.Formula);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        using var stream = new MemoryStream(); result.Value.Save(stream); stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        Assert.Equal(name + "(1,2)", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData(16)]
    [InlineData(85)]
    [InlineData(89)]
    public void A_value_aggregates_respect_required_and_destination_variadic_argument_limits(int index) {
        Assert.False(RenderFunction(index, 0).IsComplete);
        Assert.True(RenderFunction(index, 255).IsComplete);
        Assert.False(RenderFunction(index, 256).IsComplete);
    }
}
