using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(17, -2.5d, -2d, "CEILING(-2.5,-2)", -4d)]
    [InlineData(55, -2.5d, -2d, "FLOOR(-2.5,-2)", -2d)]
    [InlineData(17, 1.5d, 0d, "CEILING(1.5,0)", 0d)]
    [InlineData(55, 0d, null, "FLOOR(0,)", 0d)]
    public void Native_multiple_rounding_recalculates_signed_zero_and_omitted_operands_after_saved_conversion(
        int index, double number, double? factor, string expression, double expected) {
        var nodes = new List<byte[]> {
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, number))),
            BytesField(1, factor.HasValue ? Message(VarintField(1, 17), DoubleField(4, factor.Value))
                : Message(VarintField(1, 22))),
            BytesField(1, Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, 2)))
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
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        IWorkFormulaCacheAssertions.HasType(reopened, 0, "A1", "n");
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
