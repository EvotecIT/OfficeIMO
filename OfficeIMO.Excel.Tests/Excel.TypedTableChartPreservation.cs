using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelTypedTableChartPreservationTests {
    [Theory]
    [InlineData("Named")]
    [InlineData("Selector")]
    [InlineData("Shared")]
    [InlineData("Native")]
    [InlineData("NativeRename")]
    [InlineData("NativeTitle")]
    [InlineData("CustomProperty")]
    [InlineData("CustomPropertyCollection")]
    [InlineData("Rename")]
    public void AddingChartAfterTypedTablePreservesSavedRows(string route) {
        var orders = new[] {
            new Order { Customer = "Northwind", Total = 12.5m },
            new Order { Customer = "Contoso", Total = 19.75m }
        };
        using var output = new MemoryStream();
        using (var book = ExcelDocument.Create()) {
            ExcelSheet sheet = book.AddSheet("Orders");
            sheet.AddTable(orders);
            switch (route) {
                case "Named":
                    sheet.AddChart(OfficeChartType.Column, orders, nameof(Order.Customer), nameof(Order.Total));
                    break;
                case "Selector":
                    sheet.AddChart(OfficeChartType.Column, orders, item => item.Customer,
                        OfficeSeries.From(orders, item => item.Total));
                    break;
                case "Shared":
                    sheet.AddChart(OfficeChartType.Column, new OfficeChartData(new[] { "Northwind", "Contoso" },
                        new[] { new OfficeSeries("Total", new[] { 12.5, 19.75 }) }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
                    break;
                default:
                    ExcelChart chart = sheet.AddChart(OfficeIMO.OfficeChartType.ColumnClustered, new ExcelChartData(new[] { "Northwind", "Contoso" },
                        new[] { new ExcelChartSeries("Total", new[] { 12.5, 19.75 }) }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
                    if (route == "NativeRename") chart.Name = "Order totals";
                    if (route == "NativeTitle") chart.SetTitle("Order totals");
                    if (route == "CustomProperty") book.SetCustomDocumentProperty("Review", "Complete");
                    if (route == "CustomPropertyCollection") book.CustomDocumentProperties.Add("Review",
                        new ExcelCustomProperty { Value = "Complete", PropertyType = ExcelCustomPropertyType.Text });
                    if (route == "Rename") {
                        chart.Name = "Order totals";
                        sheet.Name = "Renamed orders";
                    }
                    break;
            }
            book.Save(output);
        }
        using var saved = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        var workbook = saved.WorkbookPart!;
        var ordersSheet = workbook.Workbook.Sheets!.Elements<Sheet>().Single(sheet =>
            sheet.Name == (route == "Rename" ? "Renamed orders" : "Orders"));
        Assert.Equal(2, workbook.Workbook.Sheets!.Elements<Sheet>().Count());
        if (route.StartsWith("CustomProperty", StringComparison.Ordinal))
            Assert.Contains("Complete", saved.CustomFilePropertiesPart!.Properties!.InnerText);
        var part = (WorksheetPart)workbook.GetPartById(ordersSheet.Id!);
        Assert.Single(part.TableDefinitionParts);
        var chartPart = Assert.Single(part.DrawingsPart!.ChartParts);
        var categories = Assert.Single(chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData>());
        var cache = Assert.Single(categories.Descendants<DocumentFormat.OpenXml.Drawing.Charts.StringCache>());
        Assert.Equal(new[] { "Northwind", "Contoso" }, cache.Elements<DocumentFormat.OpenXml.Drawing.Charts.StringPoint>()
            .Select(point => point.NumericValue!.Text));
        var rows = part.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
        Assert.Equal(3, rows.Length);
        Assert.Equal(new[] { "12.5", "19.75" }, rows.Skip(1).Select(row =>
            row.Elements<Cell>().Single(cell => cell.CellReference!.Value!.StartsWith("B", StringComparison.Ordinal)).CellValue!.Text));
        string Text(Cell cell) => cell.DataType?.Value == CellValues.SharedString
            ? workbook.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue!.Text)).InnerText
            : cell.InlineString?.InnerText ?? cell.InnerText;
        Assert.Equal(new[] { "Customer", "Northwind", "Contoso" }, rows.Select(row => Text(row.Elements<Cell>().First())));
    }

    public sealed class Order {
        public string Customer { get; set; } = "";
        public decimal Total { get; set; }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WorksheetCopyOrRemovalObservesRowsRetainedByNativeChart(bool remove) {
        using var book = ExcelDocument.Create();
        ExcelSheet sheet = book.AddSheet("Orders");
        sheet.AddTable(new[] { new Order { Customer = "Northwind", Total = 12.5m } });
        sheet.AddChart(OfficeIMO.OfficeChartType.ColumnClustered, new ExcelChartData(new[] { "Northwind" },
            new[] { new ExcelChartSeries("Total", new[] { 12.5 }) }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 4 });
        if (remove) book.RemoveWorksheet("Orders");
        else book.CopyWorksheet(sheet, "Copied orders");
        using var saved = SpreadsheetDocument.Open(new MemoryStream(book.ToBytes()), false);
        var workbook = saved.WorkbookPart!;
        var sheets = workbook.Workbook.Sheets!.Elements<Sheet>().ToArray();
        if (remove) {
            Assert.DoesNotContain(sheets, item => item.Name == "Orders");
        } else {
            foreach (string name in new[] { "Orders", "Copied orders" }) {
                var native = sheets.Single(item => item.Name == name);
                var part = (WorksheetPart)workbook.GetPartById(native.Id!);
                Assert.Equal(2, part.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Count());
                Assert.Single(part.TableDefinitionParts);
            }
        }
    }
}
