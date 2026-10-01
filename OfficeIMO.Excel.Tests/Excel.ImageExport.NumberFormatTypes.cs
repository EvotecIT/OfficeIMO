using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class ExcelImageExportTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Image_number_formats_preserve_text_and_formula_text_types(bool formula, bool merged) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddWorksheet("Text");
        if (formula) sheet.CellFormulaWithTextCache(1, 1, "=\"-5\"", "-5");
        else sheet.CellAt(1, 1).SetValue("-5");
        sheet.CellAt(1, 1).SetNumberFormat("0;[Red]0;0;@").SetFontColor("0000FF");
        if (merged) sheet.Range("A1:B1").Merge();
        using var saved = new MemoryStream(); workbook.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        // Selecting B1 exercises recovery of a merged origin outside the range.
        ExcelRange range = reopened.Sheets[0].Range(merged ? "B1:B1" : "A1:A1");
        ExcelVisualCell cell = Assert.Single(range.CreateVisualSnapshot().Cells, c => c.Column == 1);
        Assert.Equal("-5", cell.Text);
        Assert.Equal(ExcelVisualCellValueKind.Text, cell.ValueKind);
        if (formula) Assert.IsType<string>(reopened.Sheets[0].CellAt(1, 1).GetValue().Value);
        Assert.Contains(">-5</text>", range.ToSvg(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("\"$\"0.00", "$1234.50")]
    [InlineData("\"$\"#,##0.00", "$1,234.50")]
    public void Image_currency_grouping_uses_the_format_placeholders(string format, string expected) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddWorksheet("Numeric");
        sheet.CellAt(1, 1).SetValue(1234.5d).SetNumberFormat(format);
        sheet.CellAt(2, 1).SetValue(1234.5d);
        sheet.CellFormula(2, 1, "=1234.5");
        sheet.CellAt(2, 1).SetNumberFormat(format);
        using var saved = new MemoryStream(); workbook.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelRange range = reopened.Sheets[0].Range("A1:A2");
        Assert.All(range.CreateVisualSnapshot().Cells, cell => {
            Assert.Equal(expected, cell.Text);
            Assert.Equal(ExcelVisualCellValueKind.Number, cell.ValueKind);
        });
        Assert.Contains(">" + expected + "</text>", range.ToSvg(), StringComparison.Ordinal);
    }
}
