---
title: Convert CSV to Excel
description: Import CSV files into Excel worksheets as typed tables and export worksheets back to CSV with OfficeIMO.Excel.Csv, PSWriteOffice, and the officeimo CLI.
order: 54
packages:
  - OfficeIMO.Excel.Csv
  - OfficeIMO.CSV
limits: "Conversion reports list omitted headers and schema projection; set `FailOnLoss` in the export options to reject them before writing."
---

`OfficeIMO.Excel.Csv` connects the CSV and Excel packages. CSV parsing stays in `OfficeIMO.CSV` and workbook writing stays in `OfficeIMO.Excel`, so neither package depends on the other unless you install the adapter.

```powershell
dotnet add package OfficeIMO.Excel.Csv
```

## Import a CSV file as an Excel table

By default the import detects the delimiter, infers column types, and creates a filterable table, so numbers and dates arrive as numbers and dates rather than text.

### C#

```csharp
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;

using var book = ExcelDocument.Create();
ExcelCsvImportResult imported = book.ImportCsvFile("sales.csv", new ExcelCsvImportOptions {
    SheetName = "Sales",
    TableName = "Sales",
    ReaderOptions = new CsvReaderOptions { DetectDelimiter = true, InferSchema = true, SchemaSampleRows = 1000 }
});
Console.WriteLine($"Delimiter '{imported.Delimiter}', table {imported.TableName}");
book.Save("sales.xlsx");
```

### PowerShell

```powershell
Import-OfficeCsv -Path .\sales.csv -InferSchema | Export-OfficeExcel -Path .\sales.xlsx -Sheet Sales -TableName Sales
```

### CLI

```sh
officeimo tabular convert sales.csv sales.xlsx
officeimo tabular convert sales-pl.csv sales.xlsx --delimiter ';'
```

## Save a loaded CSV document as Excel

### C#

```csharp
using OfficeIMO.CSV;
using OfficeIMO.Excel.Csv;

CsvDocument csv = CsvDocument.Load("sales.csv");
csv.SaveAsExcel("sales-document.xlsx");
```

## Export a worksheet to CSV

### C#

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;

using var book = ExcelDocument.Load("sales.xlsx");
book["Sales"].SaveAsCsv("sales-export.csv");
```

### PowerShell

```powershell
Import-OfficeExcel -Path .\sales.xlsx -Sheet Sales | Export-OfficeCsv -Path .\sales-export.csv
```

### CLI

```sh
officeimo tabular convert sales.xlsx sales-export.csv --sheet Sales
```
