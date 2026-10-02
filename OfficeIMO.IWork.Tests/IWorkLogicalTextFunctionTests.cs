using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void True_function_preserves_a_stale_boolean_cache_and_recalculates_a_typed_true_result() {
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, 0d,
            boolean: true, hasFormula: true, formulaPayload: FunctionPayload(156, 0)) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell source = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.True(source.FormulaIsComplete);
        Assert.Equal("=TRUE()", source.Formula);
        Assert.Equal(false, source.Value);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal("TRUE()", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.False(reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", "b");
        Assert.Equal(1, reopened.Calculate());
        Assert.True(reopened.Sheets[0].CellAt(1, 1).GetValue<bool>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", "b");
    }

    [Theory]
    [InlineData(82, "LOWER", "Mi\"XeD", "mi\"xed")]
    [InlineData(158, "UPPER", "Mi\"XeD", "MI\"XED")]
    [InlineData(155, "TRIM", "    a    b   ", "a b")]
    [InlineData(155, "TRIM", "", "")]
    public void Text_function_literals_preserve_quotes_and_typed_saved_results(int index, string name, string input, string expected) {
        byte[] formula = BytesField(1, Message(
            BytesField(1, Message(VarintField(1, 19), StringField(6, input))),
            BytesField(1, Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, 1)))));
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, 0d,
            textValue: expected, hasFormula: true, formulaPayload: formula) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell source = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.True(source.FormulaIsComplete);
        string expression = name + "(\"" + input.Replace("\"", "\"\"") + "\")";
        Assert.Equal("=" + expression, source.Formula);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(expression, reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", "str");
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(49, 2)]
    [InlineData(69, 1)]
    [InlineData(70, 1)]
    [InlineData(304, 1)]
    [InlineData(305, 1)]
    [InlineData(82, 1)]
    [InlineData(96, 1)]
    [InlineData(155, 1)]
    [InlineData(156, 0)]
    [InlineData(158, 1)]
    public void Logical_and_text_functions_require_their_exact_argument_counts(int index, int count) {
        Assert.True(RenderFunction(index, count).IsComplete);
        if (count > 0) Assert.False(RenderFunction(index, count - 1).IsComplete);
        Assert.False(RenderFunction(index, count + 1).IsComplete);
    }
}

internal static class IWorkFormulaCacheAssertions {
    internal static void HasType(ExcelDocument document, int sheetIndex, string address, string expectedType) {
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var package = new ZipArchive(saved, ZipArchiveMode.Read);
        using Stream xml = package.GetEntry($"xl/worksheets/sheet{sheetIndex + 1}.xml")!.Open();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XElement cell = XDocument.Load(xml).Descendants(ns + "c").Single(cell => cell.Attribute("r")?.Value == address);
        Assert.NotNull(cell.Element(ns + "f"));
        Assert.Equal(expectedType, cell.Attribute("t")?.Value);
    }
}
