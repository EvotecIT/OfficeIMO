using System.IO;
using System.Linq;
using OfficeIMO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelWorksheetTableInspectionTests {
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
