using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Explicit_empty_text_and_formula_text_cache_remain_text_after_XLSX_save(bool formula) {
        byte[] expression = BytesField(1, Message(BytesField(1,
            Message(VarintField(1, 19), StringField(6, string.Empty)))));
        using var package = CreateNumbersPackage(new[] {
            new TableSpec("Empty text", 1, 1, 0d, textValue: string.Empty,
                hasFormula: formula, formulaPayload: formula ? expression : null)
        });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        var source = Assert.Single(result.Projection.Sheets.Single().Tables.Single().Cells);
        Assert.Equal(formula ? IWorkCellKind.Formula : IWorkCellKind.Text, source.Kind);
        Assert.Equal(IWorkCellKind.Text, source.ValueKind);
        Assert.Equal(string.Empty, source.Value);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var sheet = reopened.Sheets.Single();
        Assert.True(sheet.TryGetCellValueSnapshot(1, 1, out var value));
        Assert.Equal(formula ? ExcelCellValueKind.Formula : ExcelCellValueKind.Text, value!.Kind);
        Assert.Equal(string.Empty, sheet.CellAt(1, 1).GetValue<string>());
        if (formula) Assert.Equal("\"\"", sheet.GetFormulaText(1, 1));
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
