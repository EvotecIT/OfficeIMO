using System.Threading;
using System.Data;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public sealed class ExcelTableRecordReadingTests {
    [Fact]
    public void SheetAndTableReadPositionalRecordsAndExcludeTableTotalsAfterReload() {
        var expected = new[] { new Order(1, "North", 12.50m), new Order(2, "South", 25m) };
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            ExcelSheet sheet = document.AddSheet("Orders");
            ExcelTable table = sheet.AddTable(expected, new OfficeTableOptions { Name = "Orders" });
            Assert.Equal(expected, sheet.RowsAs<Order>().ToArray());
            Assert.Equal(expected, table.RowsAs<Order>().ToArray());
            Assert.Equal(expected, sheet.RowsAsParallel<Order>().ToArray());
            table.Resize("A1:C4"); // Reserve the final row for native table totals.
            table.SetTotals(new Dictionary<string, ExcelTableTotalsFunction> {
                [nameof(Order.Total)] = ExcelTableTotalsFunction.Sum
            });
            sheet.CellValue(10, 1, "Outside the table");
            Assert.Equal(expected, table.RowsAs<Order>().ToArray());
            document.Save(stream);
        }
        stream.Position = 0;
        using var loaded = ExcelDocument.Load(stream);
        Assert.Equal(expected, loaded.GetSheet("Orders").Table("Orders").RowsAs<Order>().ToArray());
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void HeaderlessTablesUseTheirNativeSchemaAndKeepTheFirstDataRow(bool dataTable, bool projected, bool readFirst) {
        var expected = new[] { new Order(1, "First", 10m), new Order(2, "Second", 20m) };
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        var options = new OfficeTableOptions { Name = "HiddenHeaders", IncludeHeaders = false };
        if (projected) {
            options.HeaderRenames[nameof(Order.Customer)] = "Customer name";
            options.Order.Add(nameof(Order.Total));
            options.Order.Add(nameof(Order.Id));
        }
        using var source = new DataTable("Source");
        source.Columns.Add(nameof(Order.Id), typeof(int));
        source.Columns.Add(nameof(Order.Customer), typeof(string));
        source.Columns.Add(nameof(Order.Total), typeof(decimal));
        foreach (Order row in expected) source.Rows.Add(row.Id, row.Customer, row.Total);
        ExcelTable table = dataTable ? sheet.AddTable(source, options) : sheet.AddTable(expected, options);
        if (readFirst) Assert.Equal(expected, table.RowsAs<Order>().ToArray());
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using var loaded = ExcelDocument.Load(stream);
        Assert.Equal(expected, loaded.GetSheet("Data").Table("HiddenHeaders").RowsAs<Order>().ToArray());
    }

    [Fact]
    public void TableReadsObserveBothCancellationSourcesWithoutClosingTheWorkbook() {
        using var document = ExcelDocument.Create();
        ExcelTable table = document.AddSheet("Data").AddTable(new[] { new Order(1, "North", 10m) });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => table.RowsAs<Order>(cancellationToken: cancellation.Token).ToArray());
        Assert.ThrowsAny<OperationCanceledException>(() => table.RowsAs<Order>(
            new ExcelReaderOptions { CancellationToken = cancellation.Token }).ToArray());
        Assert.Single(table.RowsAs<Order>());
        table.Sheet.CellValue(5, 1, "Workbook remains editable");
    }

    [Fact]
    public void StoppingTableEnumerationLeavesTheWorkbookOpenAndReadersReusable() {
        using var document = ExcelDocument.Create();
        ExcelTable table = document.AddSheet("Data").AddTable(new[] {
            new Order(1, "First", 10m), new Order(2, "Second", 20m)
        });
        using (var rows = table.RowsAs<Order>().GetEnumerator()) {
            Assert.True(rows.MoveNext());
            Assert.Equal(1, rows.Current.Id);
        }
        Assert.Equal(2, table.RowsAs<Order>().Count());
        using var saved = new MemoryStream();
        document.Save(saved);
        Assert.True(saved.Length > 0);
    }

    public sealed record Order(int Id, [property: DataColumn("Customer name")] string Customer, decimal Total);
}
