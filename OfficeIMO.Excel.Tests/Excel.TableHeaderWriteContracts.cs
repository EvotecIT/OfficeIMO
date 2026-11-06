using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, ExcelExecutionMode.Sequential)]
    [InlineData(false, ExcelExecutionMode.Parallel)]
    [InlineData(true, ExcelExecutionMode.Sequential)]
    [InlineData(true, ExcelExecutionMode.Parallel)]
    public void Test_CellWriteOwnership_TabularReplacementUsesTheFinalHeaderSchema(bool typed, ExcelExecutionMode mode) {
        using var document = ExcelDocument.Create();
        document.Execution.Mode = mode;
        document.Execution.MaxDegreeOfParallelism = 2;
        var sheet = CreateHeaderReplacementTable(document);
        if (typed) {
            sheet.WriteData(new[] { (Territory: "US", Net: 20) }, true, 1,
                ("Territory", item => (object?)item.Territory), ("Net", item => (object?)item.Net));
        } else {
            using var data = new DataTable();
            data.Columns.Add("Territory", typeof(string));
            data.Columns.Add("Net", typeof(int));
            data.Rows.Add("US", 20);
            sheet.WriteData(data, mode: mode);
        }
        Assert.Equal(new[] { "Territory", "Net" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
        Assert.Equal("Territory", sheet.Cell(1, 1).GetValue<string>());
        Assert.Equal("Net", sheet.Cell(1, 2).GetValue<string>());
        Assert.Equal("US", sheet.Cell(2, 1).GetValue<string>());
        Assert.Equal(20d, sheet.Cell(2, 2).GetValue<double>());
        Assert.Equal("SUM(Sales[Net])", Assert.Single(document["Summary"].GetFormulaCells()).Formula);
        using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(new[] { "Territory", "Net" }, Assert.Single(loaded.GetTables()).Columns.Select(column => column.Name));
        Assert.Equal("SUM(Sales[Net])", Assert.Single(loaded["Summary"].GetFormulaCells()).Formula);
        Assert.Equal(20d, loaded["Data"].Cell(2, 2).GetValue<double>());
    }

    [Theory]
    [InlineData(ExcelExecutionMode.Sequential)]
    [InlineData(ExcelExecutionMode.Parallel)]
    public void Test_CellWriteOwnership_InvalidTabularHeadersLeaveCellsAndSchemaUnchanged(ExcelExecutionMode mode) {
        using var document = ExcelDocument.Create();
        document.Execution.Mode = mode;
        document.Execution.MaxDegreeOfParallelism = 2;
        var sheet = CreateHeaderReplacementTable(document);
        using var data = new DataTable();
        data.Columns.Add(" Net ", typeof(string));
        data.Columns.Add("Net", typeof(int));
        data.Rows.Add("US", 20);
        Assert.Throws<ArgumentException>(() => sheet.WriteData(data, mode: mode));
        Assert.Equal(new[] { "Region", "Amount" }, Assert.Single(document.GetTables()).Columns.Select(column => column.Name));
        Assert.Equal("Region", sheet.Cell(1, 1).GetValue<string>());
        Assert.Equal("EU", sheet.Cell(2, 1).GetValue<string>());
        Assert.Equal(10d, sheet.Cell(2, 2).GetValue<double>());
        Assert.Equal("SUM(Sales[Amount])", Assert.Single(document["Summary"].GetFormulaCells()).Formula);
    }

    [Theory]
    [InlineData("cell")]
    [InlineData("schema")]
    [InlineData("rich")]
    [InlineData("remove")]
    public void Test_CellWriteOwnership_ColumnReferencesFollowBothImportedTableAliases(string route) {
        string path = Path.Combine(_directoryWithFiles, "HeaderAliases.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            var sheet = CreateHeaderReplacementTable(document);
            document["Summary"].Cell(2, 1).SetFormula("SUM(InternalSales[Amount])");
            document.Save();
        }
        using (var package = SpreadsheetDocument.Open(path, true)) {
            var table = package.WorkbookPart!.WorksheetParts.SelectMany(part => part.TableDefinitionParts).Single().Table!;
            table.Name = "InternalSales";
            table.DisplayName = "Sales";
            table.Save();
        }
        using var loaded = ExcelDocument.Load(path);
        var other = loaded.AddSheet("Other");
        other.Cell(1, 1).SetValue("Item");
        other.Cell(2, 1).SetValue("value");
        other.AddTable("A1:A2", name: "Sales");
        Assert.Equal("Sales2", Assert.Single(other.GetTables()).Name);
        var source = loaded["Data"];
        if (route == "cell") source.Cell(1, 2).SetValue("Net");
        else if (route == "schema") source.SetTableSchema("InternalSales", new[] { "Region", "Net" });
        else if (route == "rich") source.Cell(1, 2).SetRichText(new ExcelRichTextRun("Net"));
        else source.ResizeTable("InternalSales", "A1:A2");
        string[] expected = route == "remove" ? new[] { "SUM(#REF!)", "SUM(#REF!)" }
            : new[] { "SUM(Sales[Net])", "SUM(InternalSales[Net])" };
        Assert.Equal(expected, loaded["Summary"].GetFormulaCells().Select(item => item.Formula));
        using var reopened = ExcelDocument.Load(new MemoryStream(loaded.ToBytes()));
        Assert.Equal(expected, reopened["Summary"].GetFormulaCells().Select(item => item.Formula));
        Assert.Equal("InternalSales", Assert.Single(reopened["Data"].GetTables()).Name);
        Assert.Equal("Sales", Assert.Single(reopened["Data"].GetTables()).DisplayName);
    }

    private static ExcelSheet CreateHeaderReplacementTable(ExcelDocument document) {
        var sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Region");
        sheet.Cell(1, 2).SetValue("Amount");
        sheet.Cell(2, 1).SetValue("EU");
        sheet.Cell(2, 2).SetValue(10);
        sheet.AddTable("A1:B2", name: "Sales");
        document.AddSheet("Summary").Cell(1, 1).SetFormula("SUM(Sales[Amount])");
        return sheet;
    }
}
