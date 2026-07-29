---
title: CSV and Excel Reader
description: Read CSV, TSV, XLSX, XLSM, and XLSB through one forward-only typed API.
order: 45
---

`OfficeIMO.Tabular` is the canonical read-only API for tabular files. The same
`TabularReader` entry point handles CSV, TSV, XLSX, XLSM, and XLSB, discovers
the used range, and exposes rows through the standard .NET `DbDataReader`
contract.

Use `OfficeIMO.CSV.CsvDocument` or `OfficeIMO.Excel.ExcelDocument` when you
need to edit and save a document. Use `TabularReader` when you need to scan
rows, retrieve typed values, load a `DataTable`, or bind records.

## Install

```bash
dotnet add package OfficeIMO.Tabular
```

## Read CSV or Excel

```csharp
using OfficeIMO.Tabular;

using var reader = TabularReader.Open("sales.xlsx");
int orderId = reader.GetOrdinal("Order ID");
int total = reader.GetOrdinal("Total");

while (reader.Read()) {
    Console.WriteLine(
        $"{reader.GetInt32(orderId)}: {reader.GetDecimal(total)}");
}
```

Changing the path to `sales.csv` does not change the row-reading code.
`TabularReader` derives from `DbDataReader`, so it also works with
`DataTable.Load` and other ADO.NET consumers.

## Read workbook sheets

The first worksheet is active after `Open`. Call `NextResult()` to move
through the remaining sheets in workbook order:

```csharp
using var reader = TabularReader.Open("book.xlsx");
do {
    Console.WriteLine(reader.TableName);
    while (reader.Read()) {
        // Process this worksheet.
    }
} while (reader.NextResult());
```

Set `TabularReadOptions.TableName` when you only need one named worksheet.
Callers do not need to guess or precompute an `A1` range.

## Bind records

```csharp
using System.Runtime.Serialization;
using OfficeIMO.Tabular;

using var reader = TabularReader.Open("sales.csv");
foreach (SalesRecord sale in reader.ReadRecords<SalesRecord>()) {
    Console.WriteLine($"{sale.OrderId}: {sale.Total}");
}

[DataContract]
public sealed class SalesRecord {
    [DataMember(Name = "Order ID")]
    public int OrderId { get; set; }

    public decimal Total { get; set; }
}
```

Properties match headers case-insensitively. Use
`DataMemberAttribute.Name` when a header is not a CLR property name.
Record binding creates setters at runtime; NativeAOT and fully trimmed
applications should use the typed `DbDataReader` getters.

## Options and limits

`TabularReadOptions` controls headers, delimiter detection, type inference,
culture, spreadsheet date and numeric handling, cancellation, and the maximum
accepted input size. Streams remain owned by the caller.

The reader is forward-only and read-only. Backend CSV and Excel reader types
are intentionally not public entry points.

See the [API reference](/api/tabular/) for the complete contract and the
[benchmark evidence](/benchmarks/) for platform-specific measurements.
