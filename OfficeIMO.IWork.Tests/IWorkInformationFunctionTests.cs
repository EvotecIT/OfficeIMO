using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(69, "ISBLANK", "empty", false)]
    [InlineData(70, "ISERROR", "error", true)]
    [InlineData(70, "ISERROR", "text", false)]
    [InlineData(304, "ISNUMBER", "number", true)]
    [InlineData(304, "ISNUMBER", "text", false)]
    [InlineData(304, "ISNUMBER", "boolean", false)]
    [InlineData(305, "ISTEXT", "text", true)]
    [InlineData(305, "ISTEXT", "number", false)]
    public void Information_functions_preserve_stale_caches_and_recalculate_typed_boolean_results(
        int index, string name, string operand, bool expected) {
        byte[][] nodes = operand switch {
            "error" => new[] { Message(VarintField(1, 17), DoubleField(4, 1)),
                Message(VarintField(1, 17), DoubleField(4, 0)), VarintField(1, 4) },
            "number" => new[] { Message(VarintField(1, 17), DoubleField(4, 123)) },
            "boolean" => new[] { Message(VarintField(1, 18), VarintField(5, 1)) },
            _ => new[] { Message(VarintField(1, 19), StringField(6, operand == "empty" ? "" : "123")) }
        };
        byte[] formula = SingleCellFormula(nodes.Concat(new[] {
            Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, 1)) }).ToArray());
        using var package = CreateNumbersPackage(new[] { new TableSpec("Information", 1, 1, expected ? 0d : 1d,
            boolean: true, hasFormula: true, formulaPayload: formula) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        string argument = operand switch { "error" => "1/0", "number" => "123", "boolean" => "TRUE", "empty" => "\"\"", _ => "\"123\"" };
        Assert.Equal(name + "(" + argument + ")", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(!expected, reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", "b");
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Theory]
    [InlineData(69, "ISBLANK", false, false, true)]
    [InlineData(304, "ISNUMBER", true, false, false)]
    [InlineData(305, "ISTEXT", false, true, false)]
    public void Information_functions_follow_cross_table_type_edits_instead_of_cached_values(
        int index, string name, bool numberResult, bool textResult, bool blankResult) {
        using var package = CrossBindingPackage(formulaPayload: SingleCellFormula(SingleCellReference(true, true),
            Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, 1))));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        string target = result.WorksheetMappings[1].DestinationName.Replace("'", "''");
        Assert.Equal(name + "('" + target + "'!$A$1)", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        reopened.Sheets[1].CellValue(1, 1, 123d);
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal(numberResult, reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        reopened.Sheets[1].CellValue(1, 1, "123");
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal(textResult, reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        reopened.Sheets[1].CellAt(1, 1).Clear();
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal(blankResult, reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", "b");
    }

    [Fact]
    public void Concatenate_preserves_literal_text_and_its_qualified_argument_bounds() {
        const string first = "A\"😀";
        byte[] formula = SingleCellFormula(Message(VarintField(1, 19), StringField(6, first)),
            Message(VarintField(1, 19), StringField(6, " B")),
            Message(VarintField(1, 16), VarintField(2, 25), VarintField(3, 2)));
        using var package = CreateNumbersPackage(new[] { new TableSpec("Text", 1, 1, 0d,
            textValue: "old", hasFormula: true, formulaPayload: formula) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal("CONCATENATE(\"A\"\"😀\",\" B\")", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal("old", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(first + " B", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", "str");
        Assert.False(RenderFunction(25, 0).IsComplete);
        Assert.True(RenderFunction(25, 1).IsComplete);
        Assert.True(RenderFunction(25, 255).IsComplete);
        Assert.False(RenderFunction(25, 256).IsComplete);
    }
}
