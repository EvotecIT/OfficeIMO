#region QuickStart
using OfficeIMO;
using OfficeIMO.Excel;

Order[] orders = [
    new("Contoso", "North", 12_400m, 0.21),
    new("Fabrikam", "South", 8_900m, 0.17),
    new("Northwind", "North", 15_100m, 0.24),
    new("Tailspin", "West", 6_300m, 0.12)
];

using var book = ExcelDocument.Create("Orders.xlsx", new ExcelCreateOptions { Theme = OfficeTheme.Modern });
ExcelSheet sheet = book.AddSheet("Orders");

ExcelTable table = sheet.AddTable(orders, new OfficeTableOptions { Name = "Orders" });
sheet.AddChart(OfficeChartType.Column, orders, nameof(Order.Customer),
    new ExcelChartOptions { Title = "Revenue by customer", Row = 2, Column = 7 }, nameof(Order.Total));

book.Save();

record Order(string Customer, string Region, decimal Total, double Margin);
#endregion
