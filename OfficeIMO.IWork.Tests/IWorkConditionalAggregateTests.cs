using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(231, "SUMIFS", 3, 255, 2)]
    [InlineData(232, "AVERAGEIFS", 3, 255, 2)]
    [InlineData(233, "COUNTIFS", 2, 254, 2)]
    [InlineData(234, "AVERAGEIF", 2, 3, 1)]
    public void Conditional_aggregate_identities_accept_complete_criteria_pairs_only(
        int index, string name, int minimum, int maximum, int step) {
        foreach (int count in new[] { minimum, maximum }) {
            var formula = RenderFunction(index, count);
            Assert.True(formula.IsComplete);
            Assert.Equal("=" + name + "(" + string.Join(",", Enumerable.Range(1, count)) + ")", formula.Text);
        }
        foreach (int count in new[] { minimum - 1, maximum + 1 })
            Assert.False(RenderFunction(index, count).IsComplete);
        if (step == 2) Assert.False(RenderFunction(index, minimum + 1).IsComplete);
    }

    [Theory]
    [InlineData(231, 2)]
    [InlineData(231, 4)]
    [InlineData(231, 256)]
    [InlineData(232, 2)]
    [InlineData(232, 4)]
    [InlineData(232, 256)]
    [InlineData(233, 1)]
    [InlineData(233, 3)]
    [InlineData(233, 255)]
    [InlineData(234, 1)]
    [InlineData(234, 4)]
    public void Incomplete_conditional_aggregate_calls_retain_typed_caches_after_saved_conversion(int index, int count) {
        using var source = CreateNumbersPackage(new[] { new TableSpec("Criteria", 1, 1, 42d,
            hasFormula: true, formulaPayload: FunctionPayload(index, count)) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(source);
        var cell = Assert.Single(Assert.Single(Assert.Single(result.Projection.Sheets).Tables).Cells);
        Assert.False(cell.FormulaIsComplete);
        Assert.True(cell.CachedValueIsComplete);
        Assert.Equal(42d, cell.Value);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        using var bytes = new MemoryStream(); result.Value.Save(bytes); bytes.Position = 0;
        using var saved = ExcelDocument.Load(bytes);
        Assert.Null(saved.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(42d, saved.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(0, saved.Calculate());
        Assert.Empty(saved.ValidateOpenXml());
    }
}
