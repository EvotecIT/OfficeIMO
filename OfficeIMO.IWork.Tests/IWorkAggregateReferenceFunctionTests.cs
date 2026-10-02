using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(113, "PRODUCT", 12d, 24d)]
    [InlineData(147, "SUMSQ", 25d, 29d)]
    public void Native_aggregates_preserve_cross_table_ranges_and_recalculate_numeric_cells_after_reopen(
        int index, string name, double expected, double edited) {
        byte[] last = Message(VarintField(1, 36),
            BytesField(26, Message(VarintField(1, 2), VarintField(2, 1))),
            BytesField(27, Message(VarintField(1, 2), VarintField(2, 1))), SingleCellIdentity());
        byte[] formula = SingleCellFormula(SingleCellReference(true, true), last, VarintField(1, 29),
            Message(VarintField(1, 16), VarintField(2, (ulong)index), VarintField(3, 1)));
        using var package = CrossBindingPackage(formulaPayload: formula);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        string target = result.WorksheetMappings[1].DestinationName.Replace("'", "''");
        Assert.Equal(name + "('" + target + "'!$A$1:$B$2)", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(13d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        var data = reopened.Sheets[1];
        data.CellValue(1, 1, 3d); data.CellValue(1, 2, 4d);
        data.CellValue(2, 1, "7");
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        data.CellValue(2, 2, 2d);
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal(edited, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
