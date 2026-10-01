using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Mina_native_identity_saves_an_editable_formula_and_numeric_cache() {
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, 1d,
            hasFormula: true, formulaPayload: FunctionPayload(89, 2)) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        var source = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.True(source.FormulaIsComplete);
        Assert.Equal("=MINA(1,2)", source.Formula);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        using var stream = new MemoryStream(); result.Value.Save(stream); stream.Position = 0;
        using var reopened = ExcelDocument.Load(stream);
        Assert.Equal("MINA(1,2)", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(1d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(1d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Mina_respects_the_destination_variadic_argument_limit() {
        Assert.True(RenderFunction(89, 255).IsComplete);
        Assert.False(RenderFunction(89, 256).IsComplete);
    }
}
