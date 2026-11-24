using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Notebook.Spreadsheet;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookSpreadsheetExportTests {
    [Fact]
    public void SavedTableExportsSchemaOrderTypedCellsAndLiteralFormulaText() {
        using var resource = JsonDocument.Parse("""
            {
              "schema": { "fields": [{ "name": "Label" }, { "name": "Count" }, { "name": "Identifier" }, { "name": "Fraction" }, { "name": "Scientific" }] },
              "data": [{ "Identifier": 12345678901234567, "Count": -19, "Label": "=2+3", "Fraction": 0.12345678901234567, "Scientific": -1.25e-85 }]
            }
            """);
        NotebookOutput table = NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            ["application/vnd.dataresource+json"] = resource.RootElement
        });

        using var csv = new MemoryStream();
        Assert.Equal(1, table.SaveAsCsv(csv).ItemsWritten);
        string csvText = Encoding.UTF8.GetString(csv.ToArray());
        Assert.StartsWith("Label,Count,Identifier", csvText);
        Assert.Contains("'=2+3,-19,12345678901234567,0.12345678901234567,-1.25e-85", csvText);

        var options = new CsvSaveOptions { Delimiter = ';' };
        using var formatted = new MemoryStream();
        table.SaveAsCsv(formatted, options);
        Assert.Contains("'=2+3;-19;", Encoding.UTF8.GetString(formatted.ToArray()));
        Assert.Equal(CsvFormulaInjectionPolicy.Preserve, options.FormulaInjectionPolicy);

        options.FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Preserve;
        using var literal = new MemoryStream();
        table.SaveAsCsv(literal, options);
        Assert.Contains("\n=2+3;-19;", Encoding.UTF8.GetString(literal.ToArray()));

        using var destination = new MemoryStream();
        Assert.Equal(1, table.SaveAsExcel(destination).ItemsWritten);
        destination.Position = 0;
        using ExcelDocument workbook = ExcelDocument.Load(destination);
        ExcelSheet sheet = workbook["Notebook data"];
        Assert.Equal("Label", sheet.Cell(1, 1).GetValue<string>());
        Assert.Equal("=2+3", Assert.IsType<string>(sheet.Cell(2, 1).GetValue().Value));
        Assert.Equal(-19, sheet.Cell(2, 2).GetValue<int>());
        Assert.Equal("12345678901234567", Assert.IsType<string>(sheet.Cell(2, 3).GetValue().Value));
        Assert.Equal("0.12345678901234567", Assert.IsType<string>(sheet.Cell(2, 4).GetValue().Value));
        Assert.Equal(-1.25e-85, sheet.Cell(2, 5).GetValue<double>());
    }
}
