# OfficeIMO.Excel — .NET Excel Utilities

OfficeIMO.Excel provides a lightweight, typed, and ergonomic API for reading and writing .xlsx files on top of Open XML. It focuses on fast values reads, editable row workflows, and write helpers that avoid extra file handles.

## Quick Read Patterns

These helpers streamline reading Excel without extra reader boilerplate. They reuse the open `ExcelDocument` handle and infer headers/types for you.

```csharp
using OfficeIMO.Excel;

// Open workbook
using var doc = ExcelDocument.Load(path);

// Sheet access by name or 0-based index
var s1 = doc["Data"]; // case-insensitive
var s2 = doc[0];

// Values-only: iterate rows as dictionaries (UsedRange)
foreach (var row in s1.Rows()) {
    var name = (string)row["Name"];
    var val  = Convert.ToInt32(row["Value"]);
}

// Read a specific range and map to POCOs
var people = s1.RowsAs<Person>("A1:C10").ToList();

// Editable rows: read → edit → save (first row = headers)
foreach (var row in s1.RowsObjects()) {
    if (row.Get<int>("Value") == 10) {
        row.Set("Status", "Processed");
    }
}
doc.Save();

public sealed class Person {
    public string Name { get; set; }
    public int    Value { get; set; }
    public string Status { get; set; }
}
```

## Fluent Read

```csharp
// Values as dictionaries
var rows = doc.Read()
              .Sheet("Data")
              .UsedRange()
              .NumericAsDecimal(true)
              .AsRows()
              .ToList();

// Map to POCOs
var people = doc.Read()
                .Sheet("Data")
                .Range("A1:C10")
                .AsObjects<Person>()
                .ToList();

// Editable rows
foreach (var row in doc.Read().Sheet("Data").UsedRange().AsEditableRows())
{
    if (row.Get<int>("Value") >= 100)
        row.Set("Status", "Hold");
}
```

Notes:
- `Rows()` materializes dictionaries using the first row of the range as headers.
- `RowsObjects()` returns editable row handles; setting `cell.Value` or calling `row.Set(header, value)` writes to the sheet.
- All helpers share a single open file handle; no extra opens.
- Header sugar on sheet: `sheet.SetByHeader(row, "Status", "Processed")`, `sheet.ColumnIndexByHeader("Value")`.
- Prefer decimals? Use `ExcelReadPresets.DecimalFirst()` or set `new ExcelReadOptions { NumericAsDecimal = true }`.

### Column formatting by header

Use the discoverable builder to apply formats by header:

```csharp
// using System.Globalization;
var s = doc["Data"];

// Numbers
s.ColumnStyleByHeader("Value").Number(decimals: 2);
s.ColumnStyleByHeader("Count").Integer();

// Percent & currency
s.ColumnStyleByHeader("Rate").Percent(decimals: 1);
s.ColumnStyleByHeader("Amount").Currency(decimals: 2, culture: CultureInfo.GetCultureInfo("en-US"));

// Dates & durations
s.ColumnStyleByHeader("When").Date("yyyy-mm-dd");
s.ColumnStyleByHeader("When").DateTime("yyyy-mm-dd hh:mm:ss");
s.ColumnStyleByHeader("Duration").DurationHours();

// Custom Excel number format
s.ColumnStyleByHeader("Misc").NumberFormat("0.00E+00");
```

## Status

- Values-only read: available (`Rows`, `Rows("A1:C3")`, `RowsAs<T>`)
- Editable rows: available (`RowsObjects()`)
- Fluent read/write and header-wide helpers: planned
