using OfficeIMO.Excel;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelConditionalAverageErrorTests {
    [Theory]
    [InlineData("AVERAGEIF(A1:A2,D1,B1:B2)", false)]
    [InlineData("AVERAGEIFS(B1:B2,A1:A2,D1)", false)]
    [InlineData("AVERAGEIF(A1:A2,D1,B1:B2)", true)]
    [InlineData("AVERAGEIFS(B1:B2,A1:A2,D1)", true)]
    public void Empty_conditional_averages_return_typed_errors_and_recalculate_after_edits(string formula, bool noNumbers) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Conditional averages");
        sheet.Cell(1, 1).SetValue("EU"); sheet.Cell(2, 1).SetValue("US");
        if (noNumbers) sheet.Cell(1, 2).SetValue("text");
        else sheet.Cell(1, 2).SetValue(10d);
        sheet.Cell(2, 2).SetValue(20d);
        sheet.Cell(1, 4).SetValue(noNumbers ? "EU" : "None");
        string[] formulas = { formula, "IFERROR(" + formula + ",99)",
            "IFNA(" + formula + ",99)", "ISERROR(" + formula + ")" };
        for (int row = 1; row <= formulas.Length; row++) {
            sheet.Cell(row, 5).SetValue(42d);
            sheet.Cell(row, 5).SetFormula(formulas[row - 1]);
        }
        Assert.Equal(4, document.Calculate());
        string[] expected = { "#DIV/0!", "99", "#DIV/0!", "1" };
        for (int row = 1; row <= expected.Length; row++) {
            Assert.True(sheet.TryGetCachedFormulaValue(row, 5, out string? cache));
            Assert.Equal(expected[row - 1], cache);
        }
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        var destination = Assert.Single(reopened.Sheets);
        Assert.Equal("#DIV/0!", destination.Cell(1, 5).GetValue<string>());
        Assert.Equal(99d, destination.Cell(2, 5).GetValue<double>());
        Assert.True(destination.Cell(4, 5).GetValue<bool>());
        Assert.Empty(reopened.ValidateOpenXml());
        if (noNumbers) destination.Cell(1, 2).SetValue(10d);
        else destination.Cell(1, 4).SetValue("EU");
        Assert.Equal(4, reopened.Calculate());
        for (int row = 1; row <= 3; row++) Assert.Equal(10d, destination.Cell(row, 5).GetValue<double>());
        Assert.False(destination.Cell(4, 5).GetValue<bool>());
        using var edited = new MemoryStream(); reopened.Save(edited); edited.Position = 0;
        using var final = ExcelDocument.Load(edited);
        Assert.Equal(10d, final.Sheets[0].Cell(1, 5).GetValue<double>());
        Assert.Empty(final.ValidateOpenXml());
    }

    [Fact]
    public void Empty_conditional_average_errors_and_dependent_handlers_match_fresh_native_Excel() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root,
            "numbers-conditional-aggregates-v14.5-decode.json")));
        using var document = ExcelDocument.Load(Path.Combine(root,
            manifest.RootElement.GetProperty("emptyAverageInput").GetString()!));
        var sheet = Assert.Single(document.Sheets);
        Assert.Equal(16, document.Calculate());
        Verify(sheet);
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Verify(reopened.Sheets[0]);
        Assert.Empty(reopened.ValidateOpenXml());

        void Verify(ExcelSheet destination) {
            foreach (JsonElement cell in manifest.RootElement.GetProperty("emptyAverageCells").EnumerateArray()) {
                int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
                Assert.Equal(cell.GetProperty("formula").GetString(), destination.GetFormulaText(row, column));
                Assert.True(destination.TryGetCachedFormulaValue(row, column, out string? cache));
                Assert.Equal(cell.GetProperty("value").GetString(), cache);
                if (cell.GetProperty("kind").GetString() == "b") Assert.True(destination.Cell(row, column).GetValue<bool>());
            }
        }
    }
}
