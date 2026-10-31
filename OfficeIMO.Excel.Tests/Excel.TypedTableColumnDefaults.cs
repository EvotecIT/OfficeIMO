using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelTypedTableColumnDefaultsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TypedTablesFitProjectedColumnsAndKeepExplicitWorksheetWidths(bool dataTable) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        sheet.SetColumnWidth(3, ExcelColumnWidth.Characters(6));
        sheet.SetColumnWidth(6, ExcelColumnWidth.Characters(9));
        var options = new OfficeTableOptions { Include = { "Name", "Amount" }, Order = { "Amount", "Name" } };
        options.Alignments.Add(OfficeAlignment.Center);
        ExcelTable table;
        const string name = "Northwind international sales division";
        if (dataTable) {
            using var rows = new DataTable();
            rows.Columns.Add("Name", typeof(string));
            rows.Columns.Add("Amount", typeof(decimal));
            rows.Columns.Add("Ignored", typeof(string));
            rows.Rows.Add(name, 123m, "Unused");
            table = sheet.AddTable(rows, options, row: 4, column: 3);
        } else {
            table = sheet.AddTable(new[] { new { Name = name, Amount = 123m, Ignored = "Unused" } },
                options, row: 4, column: 3);
        }
        Assert.Equal("C4:D5", table.Range);
        using var output = new MemoryStream();
        document.Save(output);
        output.Position = 0;
        using var saved = SpreadsheetDocument.Open(output, false);
        var worksheet = saved.WorkbookPart!.WorksheetParts.Single().Worksheet;
        var columns = worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
        double Width(uint index) => columns.Single(column => column.Min!.Value <= index && column.Max!.Value >= index).Width!.Value;
        Assert.Equal(6, Width(3));
        Assert.True(Width(4) > 20, "The long projected text should receive a fitted native column width.");
        Assert.Equal(9, Width(6));
        Assert.True(sheet.TryGetCellText(5, 4, out string text));
        Assert.Equal(name, text);
    }
}
