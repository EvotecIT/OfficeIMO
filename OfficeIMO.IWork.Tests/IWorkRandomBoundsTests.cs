using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(2147483648d, "RANDBETWEEN(2147483648,2147483648)")]
    [InlineData(-2147483649d, "RANDBETWEEN(-2147483649,-2147483649)")]
    [InlineData(9007199254740991d, "RANDBETWEEN(9007199254740991,9007199254740991)")]
    [InlineData(-9007199254740991d, "RANDBETWEEN(-9007199254740991,-9007199254740991)")]
    public void Native_large_random_bounds_preserve_cache_then_recalculate_after_saved_conversion(double number, string expression) {
        var nodes = new List<byte[]> {
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, number))),
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, number))),
            BytesField(1, Message(VarintField(1, 16), VarintField(2, 119), VarintField(3, 2)))
        };
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, 42d,
            hasFormula: true, formulaPayload: BytesField(1, Message(nodes.ToArray()))) });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.Equal("=" + expression, result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.Formula);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(expression, reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(1, reopened.Calculate());
        Assert.Equal(number, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", "n");
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
