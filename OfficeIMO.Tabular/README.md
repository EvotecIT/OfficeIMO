# OfficeIMO.Tabular

OfficeIMO.Tabular provides one forward-only API for reading CSV, TSV, XLSX,
XLSM, and XLSB data. It is the read-only entry point when the caller wants rows,
typed values, or objects without editing and saving the source document.

Use `OfficeIMO.CSV.CsvDocument` or `OfficeIMO.Excel.ExcelDocument` instead when
the file must be transformed, inspected, edited, or written again.

## Read rows

```csharp
using OfficeIMO.Tabular;

using var reader = TabularReader.Open("sales.xlsx");
int orderId = reader.GetOrdinal("Order ID");
int total = reader.GetOrdinal("Total");

while (reader.Read()) {
    Console.WriteLine($"{reader.GetInt32(orderId)}: {reader.GetDecimal(total)}");
}
```

The same code works for a CSV path. `TabularReader` derives from
`DbDataReader`, so it can also be passed to `DataTable.Load` and APIs that
accept an ADO.NET reader.

## Read records

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

Property names match headers case-insensitively. Use
`[DataMember(Name = "...")]` when a header is not a CLR property name.
Object binding uses runtime-generated setters; NativeAOT and fully trimmed
applications should use the typed `DbDataReader` getters instead.

## Read workbook results

Workbook sheets are exposed in workbook order. The first sheet is active after
`Open`; call `NextResult()` to move to the next one.

```csharp
using var reader = TabularReader.Open("book.xlsx");
do {
    Console.WriteLine(reader.TableName);
    while (reader.Read()) {
        // Process the current worksheet.
    }
} while (reader.NextResult());
```

Set `TabularReadOptions.TableName` to open one named worksheet. Used ranges are
discovered automatically; callers do not have to guess an `A1` range.

## Options and boundaries

`TabularReadOptions` controls headers, CSV delimiter detection, type inference,
culture, spreadsheet date and numeric handling, cancellation, and the maximum
accepted input size. Streams remain owned by the caller.

The reader is forward-only and read-only. It does not expose backend CSV or
Excel reader types, and it does not replace the editable document models.
