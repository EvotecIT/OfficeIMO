---
title: Import and Export Data
description: Move data between Excel and .NET objects, DataTable, IDataReader, and CSV with OfficeIMO.Excel, PSWriteOffice, and the officeimo CLI.
order: 27
---

OfficeIMO.Excel reads and writes the shapes .NET code already uses: typed objects, `DataTable`, and `IDataReader`. Header names map to property names, and cell values keep their types (numbers, dates, booleans) instead of becoming text.

## Write objects or a DataTable to a sheet

### C#

```csharp
using System.Data;
using OfficeIMO;
using OfficeIMO.Excel;

var table = new DataTable("Servers");
table.Columns.Add("Name", typeof(string));
table.Columns.Add("Cores", typeof(int));
table.Columns.Add("Patched", typeof(DateTime));
table.Rows.Add("app-01", 8, new DateTime(2026, 10, 2));
table.Rows.Add("db-01", 16, new DateTime(2026, 9, 28));

using var book = ExcelDocument.Create("Inventory.xlsx");
book.AddSheet("Servers").AddTable(table, new OfficeTableOptions { Name = "Servers" });
book.AddSheet("Owners").AddTable(new[] {
    new { Team = "Platform", Owner = "Adele Vance" },
    new { Team = "Data", Owner = "Alex Wilber" }
});
book.Save();
```

### PowerShell

```powershell
Get-Service | Select-Object Name, Status, StartType |
    Export-OfficeExcel -Path .\Services.xlsx -Sheet Services -TableName Services
```

## Read rows as objects or a DataTable

`RowsAs<T>` maps by header name. `ToDataTable` returns a typed `DataTable` for code that binds to one.

### C#

```csharp
using System.Data;
using OfficeIMO.Excel;

using var book = ExcelDocument.Load("Inventory.xlsx");

foreach (Server server in book["Servers"].RowsAs<Server>()) {
    Console.WriteLine($"{server.Name}: {server.Cores} cores");
}

DataTable owners = book["Owners"].ToDataTable();
Console.WriteLine($"{owners.Rows.Count} owners");

record Server(string Name, int Cores, DateTime Patched);
```

### PowerShell

```powershell
$rows = Import-OfficeExcel -Path .\Inventory.xlsx -Sheet Servers
$rows | Where-Object Cores -ge 16
```

## Stream a sheet through IDataReader

`ExcelDocument.OpenReader` returns an `ExcelWorkbookDataReader` (a `DbDataReader`), so rows can flow into `SqlBulkCopy` or any API that accepts a data reader without loading the workbook model.

### C#

```csharp
using System.Data;
using OfficeIMO.Excel;

using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader("Inventory.xlsx",
    new ExcelReaderOptions { SheetName = "Servers" });

while (reader.Read()) {
    string name = reader.GetString(reader.GetOrdinal("Name"));
    Console.WriteLine(name);
}
```

## Write a database result to Excel

Pass the `IDataReader` from a database command straight to the streaming writer. Rows are written as they are read.

### C#

```csharp
using System.Data;
using OfficeIMO.Excel;

static void ExportOrders(IDataReader orders, string path) {
    using ExcelWorkbookWriter writer = ExcelDocument.CreateWriter(path);
    writer.WriteTable("Orders", orders);
}
```

## Convert between CSV and Excel

Install `OfficeIMO.Excel.Csv` to import a CSV file as a typed table, or save a sheet as CSV. See [Convert CSV to Excel](/docs/csv/convert-to-excel/) for options.

### C#

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;

using var book = ExcelDocument.Create();
ExcelCsvImportResult imported = book.ImportCsvFile("sales.csv");
book[imported.SheetName].SaveAsCsv("sales-roundtrip.csv");
book.Save("sales.xlsx");
```

### CLI

```sh
officeimo tabular convert sales.csv sales.xlsx
officeimo tabular convert Inventory.xlsx servers.csv --sheet Servers
```

## Limits and reference

**Limits:** see the [Excel package row](/compatibility/#package-excel) and the [Excel CSV adapter row](/compatibility/#package-excel-csv).

**Reference:** [OfficeIMO.Excel API](/api/excel/) · [OfficeIMO.Excel.Csv API](/api/excel-csv/)
