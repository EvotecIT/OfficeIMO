using OfficeIMO.Excel;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelConditionalAggregateTypedErrorTests {
    [Fact]
    public void Conditional_aggregate_errors_and_typed_criteria_match_native_Excel_after_save_and_edits() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root,
            "numbers-conditional-aggregates-v14.5-decode.json")));
        using var document = ExcelDocument.Load(Path.Combine(root,
            manifest.RootElement.GetProperty("typedErrorInput").GetString()!));
        int count = manifest.RootElement.GetProperty("typedErrorFormulaCells").GetInt32();
        int calculated = document.Calculate();
        Verify(document.Sheets[0]);
        Assert.Equal(count, calculated);
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = ExcelDocument.Load(bytes);
        var destination = Assert.Single(reopened.Sheets);
        Verify(destination);
        Assert.Empty(reopened.ValidateOpenXml());

        destination.CellFormula(2, 2, "5");
        Assert.Equal(count, reopened.Calculate());
        for (int row = 2; row <= 7; row++) {
            for (int column = 7; column <= 9; column++) Assert.Equal(5d, destination.CellAt(row, column).GetValue<double>());
            Assert.False(destination.CellAt(row, 10).GetValue<bool>());
        }
        double[] mixed = { 15d, 15d, 7.5d, 7.5d, 5d, 10d };
        for (int row = 8; row <= 13; row++) Assert.Equal(mixed[row - 8], destination.CellAt(row, 7).GetValue<double>());
        using var edited = new MemoryStream(); reopened.Save(edited); edited.Position = 0;
        using var final = ExcelDocument.Load(edited);
        Assert.Equal(5d, final.Sheets[0].CellAt(2, 7).GetValue<double>());
        Assert.Empty(final.ValidateOpenXml());

        void Verify(ExcelSheet sheet) {
            foreach (JsonElement cell in manifest.RootElement.GetProperty("typedErrorCells").EnumerateArray()) {
                int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
                Assert.Equal(cell.GetProperty("formula").GetString(), sheet.GetFormulaText(row, column));
                Assert.True(sheet.TryGetCachedFormulaValue(row, column, out string? value), cell.GetProperty("address").GetString());
                Assert.Equal(cell.GetProperty("value").GetString(), value);
                if (cell.GetProperty("kind").GetString() == "b")
                    Assert.Equal(value == "1", sheet.CellAt(row, column).GetValue<bool>());
            }
        }
    }

    [Fact]
    public void Conditional_errors_do_not_mask_unqualified_ranges_criteria_or_resource_limits() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Qualification");
        sheet.CellValue(1, 1, "EU"); sheet.CellValue(2, 1, "EU");
        sheet.CellFormula(1, 2, "NA()"); sheet.CellFormula(2, 2, "UNKNOWN()");
        string[] expressions = {
            "SUMIF(A1:A2,\"EU\",B1:B2)", "AVERAGEIF(A1:A2,\"EU\",B1:B2)",
            "SUMIFS(B1:B2,A1:A2,\"EU\")", "AVERAGEIFS(B1:B2,A1:A2,\"EU\")",
            "MINIFS(B1:B2,A1:A2,\"EU\")", "MAXIFS(B1:B2,A1:A2,\"EU\")",
            "SUMIFS(B1,A1,NA(),A1,UNKNOWN())", "COUNTIFS(A1,NA(),A1,UNKNOWN())",
            "AVERAGEIFS(B1,A1,NA(),A1,UNKNOWN())", "SUMIF(A1,NA(),UNKNOWN())",
            "AVERAGEIF(A1,NA(),UNKNOWN())", "SUMIFS(B1,A1,NA(),A1:A100000,\"EU\")",
            "AVERAGEIFS(B1,A1,NA(),A1:A100000,\"EU\")", "COUNTIFS(A1,NA(),A1:A100000,\"EU\")"
        };
        for (int index = 0; index < expressions.Length; index++) {
            sheet.CellValue(index + 1, 4, 42d);
            sheet.CellFormula(index + 1, 4, "IFERROR(" + expressions[index] + ",99)");
        }
        Assert.Equal(1, document.Calculate()); // Only the independent NA operand is qualified.
        for (int row = 1; row <= expressions.Length; row++) {
            Assert.True(sheet.TryGetCachedFormulaValue(row, 4, out string? value));
            Assert.Equal("42", value);
        }
    }
}
