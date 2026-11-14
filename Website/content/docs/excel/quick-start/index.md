---
title: Excel Quick Start
description: Create an Excel workbook with a typed table and a chart, then read it back, with OfficeIMO.Excel, PSWriteOffice, or the officeimo CLI.
order: 23
packages:
  - OfficeIMO.Excel
search_terms:
  - create excel
  - read excel
  - xlsx
  - write excel file
  - excel c#
commands:
  - officeimo tabular sheets
  - officeimo tabular schema
  - officeimo tabular convert
verified: 2026-11-14
---

This page creates a workbook from a list of objects, adds a native Excel table and a chart, saves it, and reads the rows back. Nothing else needs to be installed: OfficeIMO.Excel writes the `.xlsx` package directly.

```powershell
dotnet add package OfficeIMO.Excel
```

## Create a workbook with a table and a chart

`AddTable` writes a header row from the property names, typed cells, and a table style. `AddChart` plots the same data and anchors the chart at a row and column.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

Order[] orders = [
    new("Contoso", "North", 12_400m),
    new("Fabrikam", "South", 8_900m),
    new("Northwind", "North", 15_100m)
];

using var book = ExcelDocument.Create("Orders.xlsx");
ExcelSheet sheet = book.AddSheet("Orders");

sheet.AddTable(orders, new OfficeTableOptions { Name = "Orders" });
sheet.AddChart(OfficeChartType.Column, orders, nameof(Order.Customer),
    new ExcelChartOptions { Title = "Revenue by customer", Row = 2, Column = 6 }, nameof(Order.Total));
sheet.AutoFitColumns();

book.Save();

record Order(string Customer, string Region, decimal Total);
```

### PowerShell

```powershell
$orders = @(
    [pscustomobject]@{ Customer = 'Contoso';   Region = 'North'; Total = 12400 }
    [pscustomobject]@{ Customer = 'Fabrikam';  Region = 'South'; Total = 8900 }
    [pscustomobject]@{ Customer = 'Northwind'; Region = 'North'; Total = 15100 }
)

New-OfficeExcel -Path .\Orders.xlsx {
    ExcelSheet 'Orders' {
        ExcelTable $orders -Name Orders
        ExcelChart -Table Orders -Category Customer -Value Total -Address F2 -Type Column -Title 'Revenue by customer'
    }
}
```

For a one-line export of pipeline output, use `Export-OfficeExcel`:

```powershell
Get-Process | Select-Object Name, Id, CPU | Export-OfficeExcel -Path .\Processes.xlsx -TableName Processes
```

## Read the rows back

`RowsAs<T>` maps columns to properties or constructor parameters by header name.

### C#

```csharp
using OfficeIMO.Excel;

using var book = ExcelDocument.Load("Orders.xlsx");
foreach (Order order in book["Orders"].RowsAs<Order>()) {
    Console.WriteLine($"{order.Customer}: {order.Total:N0}");
}

record Order(string Customer, string Region, decimal Total);
```

### PowerShell

```powershell
Import-OfficeExcel -Path .\Orders.xlsx -Sheet Orders | Where-Object Region -eq 'North'
```

### CLI

```sh
officeimo tabular sheets Orders.xlsx
officeimo tabular schema Orders.xlsx --sheet Orders
officeimo tabular convert Orders.xlsx orders.csv --sheet Orders
```

## Next steps

- [Format cells and numbers](/docs/excel/styles-number-formats/)
- [Write formulas](/docs/excel/formulas/)
- [Add charts](/docs/excel/charts/)
- [Import and export data](/docs/excel/import-export-data/)
- [Large files and streaming](/docs/excel/large-files/)
- [Protect sheets and workbooks](/docs/excel/protection/)
