using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(true, "text")]
    [InlineData(true, "boolean")]
    [InlineData(true, "number")]
    [InlineData(false, "error")]
    public void Iferror_preserves_stale_caches_and_recalculates_only_the_selected_typed_branch(bool primaryError, string fallback) {
        var nodes = new List<byte[]> { Message(VarintField(1, 17), DoubleField(4, 6)) };
        if (primaryError) {
            nodes.Add(Message(VarintField(1, 17), DoubleField(4, 0)));
            nodes.Add(VarintField(1, 4));
        }
        switch (fallback) {
            case "text": nodes.Add(Message(VarintField(1, 19), StringField(6, "Recovered"))); break;
            case "boolean": nodes.Add(Message(VarintField(1, 18), VarintField(5, 1))); break;
            case "number": nodes.Add(Message(VarintField(1, 17), DoubleField(4, 99))); break;
            default:
                nodes.Add(Message(VarintField(1, 17), DoubleField(4, 1)));
                nodes.Add(Message(VarintField(1, 17), DoubleField(4, 0)));
                nodes.Add(VarintField(1, 4));
                break;
        }
        nodes.Add(Message(VarintField(1, 16), VarintField(2, 235), VarintField(3, 2)));
        using var package = CreateNumbersPackage(new[] { new TableSpec("Fallback", 1, 1, 42d,
            hasFormula: true, formulaPayload: SingleCellFormula(nodes.ToArray())) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(result.IsVisualFallback);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        string alternate = fallback switch { "text" => "\"Recovered\"", "boolean" => "TRUE", "number" => "99", _ => "1/0" };
        Assert.Equal("IFERROR(" + (primaryError ? "6/0" : "6") + "," + alternate + ")", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(1, reopened.Calculate());
        if (fallback == "text") Assert.Equal("Recovered", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        else if (fallback == "boolean") Assert.True(reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        else Assert.Equal(primaryError ? 99d : 6d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", fallback == "text" ? "str" : fallback == "boolean" ? "b" : "n");
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Fact]
    public void Iferror_requires_both_the_expression_and_fallback() {
        Assert.False(RenderFunction(235, 1).IsComplete);
        Assert.True(RenderFunction(235, 2).IsComplete);
        Assert.False(RenderFunction(235, 3).IsComplete);
    }
}
