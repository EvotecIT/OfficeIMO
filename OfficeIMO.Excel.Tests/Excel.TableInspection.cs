using System.IO;
using System.Linq;
using System.Data;
using OfficeIMO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelWorksheetTableInspectionTests {
    [Fact]
    public void RemovedWorksheetDoesNotInspectReplacementWithSameName() {
        using var workbook = ExcelDocument.Create();
        var original = workbook.AddWorksheet("Orders");
        workbook.AddWorksheet("Plain");
        workbook.RemoveWorksheet("Orders");
        var replacement = workbook.AddWorksheet("Orders");
        replacement.CellValue(1, 1, "Item");
        replacement.CellValue(2, 1, "Beta");
        replacement.AddTable("A1:A2", true, "ReplacementItems", ExcelTableStyle.TableStyleMedium9);
        Assert.Throws<System.InvalidOperationException>(() => original.GetTables());
        Assert.Equal("ReplacementItems", Assert.Single(replacement.GetTables()).Name);
    }

    [Fact]
    public void SavedPackageRejectsStaleWorksheetInspectionAndFreshHandleKeepsMetadata() {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-TableInspection-" + System.Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using var workbook = ExcelDocument.Create();
            var sheet = workbook.AddWorksheet("Orders");
            sheet.CellValue(1, 1, "Item");
            sheet.CellValue(2, 1, "Alpha");
            sheet.AddTable("A1:A2", true, "OrderItems", ExcelTableStyle.TableStyleMedium9);
            workbook.Save(path);
            Assert.Throws<System.InvalidOperationException>(() => sheet.GetTables());
            Assert.Equal("OrderItems", Assert.Single(workbook.Sheets[0].GetTables()).Name);
            Assert.Equal("OrderItems", Assert.Single(workbook.GetTables()).Name);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void WorksheetSnapshotIncludesDeferredDataTableDefinition() {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Orders");
        sheet.InsertDataTableAsTable(CreateOrders(), tableName: "OrderItems");
        var table = Assert.Single(sheet.GetTables());
        Assert.Equal("OrderItems", table.Name);
        Assert.Equal("A1:B2", table.Range);
        Assert.Equal(new[] { "Item", "Count" }, table.Columns.Select(c => c.Name));
    }

    [Fact]
    public void WorkbookSnapshotIncludesDeferredDataSetDefinitions() {
        using var workbook = ExcelDocument.Create();
        using var data = new DataSet();
        data.Tables.Add(CreateOrders());
        workbook.InsertDataSet(data, createTables: true);
        var table = Assert.Single(workbook.GetTables());
        Assert.Equal("Orders", table.SheetName);
        Assert.Equal("A1:B2", table.Range);
        Assert.True(table.HasAutoFilter);
    }

    private static DataTable CreateOrders() {
        var table = new DataTable("Orders");
        table.Columns.Add("Item", typeof(string));
        table.Columns.Add("Count", typeof(int));
        table.Rows.Add("Alpha", 2);
        return table;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WorksheetSnapshotKeepsNativeMetadataAndExcludesOtherSheets(bool reopen) {
        using var source = ExcelDocument.Create();
        var orders = source.AddWorksheet("Orders");
        var stock = source.AddWorksheet("Stock");
        var plain = source.AddWorksheet("Plain");
        foreach (var sheet in new[] { orders, stock }) {
            sheet.CellValue(1, 1, "Item");
            sheet.CellValue(1, 2, "Count");
            sheet.CellValue(2, 1, "Alpha");
            sheet.CellValue(2, 2, 2);
        }
        orders.AddTable("A1:B2", true, "OrderItems", ExcelTableStyle.TableStyleMedium9);
        stock.AddTable("A1:B2", true, "StockItems", ExcelTableStyle.TableStyleMedium2);
        plain.CellValue(1, 1, "Not a named table");
        using var bytes = new MemoryStream();
        source.Save(bytes);
        using var opened = reopen ? ExcelDocument.Load(new MemoryStream(bytes.ToArray()),
            new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly }) : null;
        var workbook = opened ?? source;
        Assert.Equal(2, workbook.GetTables().Count);
        var table = Assert.Single(workbook.Sheets[0].GetTables());
        Assert.Equal("OrderItems", table.Name);
        Assert.Equal("Orders", table.SheetName);
        Assert.Equal(0, table.SheetIndex);
        Assert.Equal("A1:B2", table.Range);
        Assert.Equal("TableStyleMedium9", table.StyleName);
        Assert.True(table.HasAutoFilter);
        Assert.Equal(new[] { "Item", "Count" }, table.Columns.Select(c => c.Name));
        Assert.Equal("StockItems", Assert.Single(workbook.Sheets[1].GetTables()).Name);
        Assert.Empty(workbook.Sheets[2].GetTables());
    }
}
