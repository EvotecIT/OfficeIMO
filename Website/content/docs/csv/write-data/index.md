---
title: Write CSV from Objects and Data Readers
description: "Write CSV from objects and IDataReader results, with formula-injection protection and compression, using OfficeIMO.CSV and PSWriteOffice."
order: 52
packages:
  - OfficeIMO.CSV
cmdlets:
  - Export-OfficeCsv
search_terms:
  - write csv
  - export csv
  - objects to csv
verified: 2026-11-14
---

`CsvDocument` is the editable model for building a file in memory. For exports, write rows directly: `SaveObjects` writes a sequence of objects, and `WriteDataReader` streams a database result to disk without materializing it.

## Write objects

### C#

```csharp
using OfficeIMO.CSV;

var rows = new[] {
    new { Name = "Alpha", Count = 10, Active = true },
    new { Name = "Beta", Count = 20, Active = false }
};

CsvDocument.FromObjects(rows).Save("summary.csv");

CsvDocument.SaveObjects("summary.csv.gz", rows, new CsvSaveOptions {
    CompressionType = CsvCompressionType.Auto
});
```

The `.gz` extension selects gzip compression when `CompressionType` is `Auto`.

### PowerShell

```powershell
Get-Process | Select-Object Name, Id, CPU | Export-OfficeCsv -Path .\processes.csv
Get-Process | Select-Object Name, Id | ConvertTo-OfficeCsv -Delimiter ';'
```

## Write a database result

Pass the `IDataReader` from a database command. Rows are written as they are read.

### C#

```csharp
using System.Data;
using OfficeIMO.CSV;

static void ExportOrders(IDataReader orders) {
    CsvDocument.WriteDataReader("orders.csv", orders, new CsvSaveOptions {
        NullValue = "",
        DateTimeFormat = "yyyy-MM-ddTHH:mm:ssZ",
        UseUtc = true
    });
}
```

For very wide or CPU-heavy exports, `CsvDocument.WriteDataReaderParallel` formats row batches on several threads and writes them in the original order. Measure both on your data: for simple rows the sequential writer is often faster.

## Protect spreadsheet users from formula injection

A CSV value such as `=HYPERLINK(...)` runs as a formula when the file is opened in a spreadsheet application. Escape formula-like text when the file is meant for people.

### C#

```csharp
using OfficeIMO.CSV;

CsvDocument.Load("comments.csv").Save("comments-safe.csv", new CsvSaveOptions {
    FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
});
```

### PowerShell

```powershell
Import-OfficeCsv -Path .\comments.csv | Export-OfficeCsv -Path .\comments-safe.csv -FormulaInjectionPolicy Escape
```

## Write rows you already projected

### C#

```csharp
using OfficeIMO.CSV;

object?[][] projected = {
    new object?[] { "Alpha", 10, true },
    new object?[] { "Beta", 20, false }
};

using var output = File.CreateText("projected.csv");
using var csv = new CsvRowWriter(output);
csv.WriteRows(new[] { "Name", "Count", "Active" }, projected);
```
