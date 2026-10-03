using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void Reference_shape_scalar_literals_and_nested_current_cell_functions_recalculate_after_reopen() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Shape");
        string[] expressions = { "ROW()", "COLUMN()", "ROW()+COLUMN()", "ROWS(123)", "COLUMNS(\"A1\")", "ROWS(TRUE)", "COLUMNS(FALSE)" };
        double[] expected = { 4, 3, 8, 1, 1, 1, 1 };
        for (int i = 0; i < expressions.Length; i++) sheet.CellFormula(4, i + 2, expressions[i]);
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(expressions.Length, reopened.Calculate());
        for (int i = 0; i < expressions.Length; i++) {
            Assert.Equal(expressions[i], reopened.Sheets[0].GetFormulaText(4, i + 2));
            Assert.Equal(expected[i], reopened.Sheets[0].CellAt(4, i + 2).GetValue<double>());
        }
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
