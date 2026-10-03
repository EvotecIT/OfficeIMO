using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class ExcelImageExportTests {
    [Theory]
    [InlineData(1.5d, "0.###E+00", "1.5E+00")]
    [InlineData(1d, "0.00##E+00", "1.00E+00")]
    [InlineData(0d, "0.###E+00", "0E+00")]
    [InlineData(1250d, "0.00E-00", "1.25E03")]
    [InlineData(0.0125d, "0.00e-00", "1.25e-02")]
    [InlineData(-1250d, "0.00E+00;[Red](0.###E+00)", "(1.25E+03)")]
    [InlineData(12.5d, "\"E+\"0.00", "E+12.50")]
    [InlineData(12.5d, "0.00\"E+\"", "12.50E+")]
    [InlineData(12.5d, "\\E\\+0.00", "E+12.50")]
    [InlineData(6250d, "0.00E+000", "6.25E+003")]
    [InlineData(6.25e100d, "0.00E+00", "6.25E+100")]
    [InlineData(12.346d, "0.00\\0", "12.350")]
    [InlineData(1.5d, "\"ver.1 \"0.00##E+00", "ver.1 1.50E+00")]
    [InlineData(12.5d, "\"ver.1 \"0.00", "ver.1 12.50")]
    [InlineData(1.5d, "0.#0E+00", "1.50E+00")]
    [InlineData(1.5d, "0.#0", "1.50")]
    public void Scientific_number_formats_preserve_optional_precision_exponent_sign_and_literals(double value, string format, string expected) {
        Assert.Equal(expected, ExcelNumberFormatDisplay.FormatNumericText(value, 164U, format, value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Saved_scientific_formats_reach_numeric_cells_and_formula_caches(bool merged) {
        using var workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddWorksheet("Scientific");
        sheet.CellAt(1, 1).SetValue(-1250d).SetNumberFormat("0.00E+00;[Red](0.###E+00)");
        sheet.CellAt(2, 1).SetValue(1.5d);
        sheet.CellFormula(2, 1, "=1.5");
        sheet.CellAt(2, 1).SetNumberFormat("0.###E-00");
        sheet.CellAt(3, 1).SetValue("-1250").SetNumberFormat("0.00E+00;[Red](0.###E+00)").SetFontColor("0000FF");
        if (merged) sheet.Range("A1:B1").Merge();
        using var saved = new MemoryStream(); workbook.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelVisualCell negative = Assert.Single(reopened.Sheets[0].Range(merged ? "B1:B1" : "A1:A1").CreateVisualSnapshot().Cells, cell => cell.Column == 1);
        Assert.Equal("(1.25E+03)", negative.Text);
        Assert.Equal("FF0000", negative.Style.FontColorHex);
        ExcelRange range = reopened.Sheets[0].Range("A2:A3");
        ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
        Assert.Equal("1.5E00", Assert.Single(snapshot.Cells, cell => cell.Row == 2).Text);
        ExcelVisualCell text = Assert.Single(snapshot.Cells, cell => cell.Row == 3);
        Assert.Equal("-1250", text.Text);
        Assert.Equal("0000FF", text.Style.FontColorHex);
        Assert.Contains(">1.5E00</text>", range.ToSvg(), StringComparison.Ordinal);
    }
}
