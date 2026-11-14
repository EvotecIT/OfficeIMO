---
title: Large Files and Streaming
description: Write hundreds of thousands of Excel rows with bounded memory and read large XLSX and XLSB worksheets fast with OfficeIMO.Excel streaming writers and data readers.
order: 28
packages:
  - OfficeIMO.Excel
limits: "Benchmark results for these paths are on the [benchmarks page](/benchmarks/)."
search_terms:
  - large excel
  - streaming
  - fast excel
  - big xlsx
---

`ExcelDocument.Create` and `Load` keep the whole workbook editable until `Save()`. That is the right model for reports and templates. For exports and imports measured in hundreds of thousands of rows, use the forward-only paths below: they hold one row at a time instead of the whole sheet.

## Write a large export

`CreateWriter` returns an `ExcelWorkbookWriter`. `WriteTable` accepts `IEnumerable<T>`, `IAsyncEnumerable<T>`, a `DataTable`, or an `IDataReader`. The file is published when the writer is disposed.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

IEnumerable<LogEvent> events = Enumerable.Range(1, 200_000)
    .Select(i => new LogEvent(DateTime.UtcNow.AddSeconds(-i), i % 7 == 0 ? "Warning" : "Info", $"Event {i}"));

using (ExcelWorkbookWriter writer = ExcelDocument.CreateWriter("Events.xlsx", theme: OfficeTheme.Modern)) {
    writer.WriteTable("Events", events);
}

record LogEvent(DateTime Time, string Level, string Message);
```

## Write XLSB with a typed row callback

For the smallest files and the least work per row, write binary `.xlsb` and supply the values through a callback. Scalars go straight to native records without reflection.

### C#

```csharp
using OfficeIMO.Excel;

IEnumerable<LogEvent> events = Enumerable.Range(1, 200_000)
    .Select(i => new LogEvent(DateTime.UtcNow.AddSeconds(-i), "Info", $"Event {i}"));

using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xlsb, "Events.xlsb")) {
    writer.WriteTable("Events", events, new[] { "Time", "Level", "Message" },
        static (row, item) => row.Write(item.Time).Write(item.Level).Write(item.Message));
}

record LogEvent(DateTime Time, string Level, string Message);
```

Binary row writers create plain worksheets: tables, AutoFit, freeze panes, and declared styles are rejected. XLS (`ExcelFileFormat.Xls`) is limited to 65,536 rows and 256 columns.

## Read a large worksheet

`OpenReader` returns a forward-only `DbDataReader`. Map rows to objects with `RowsAs<T>` or read fields directly.

### C#

```csharp
using OfficeIMO.Data;
using OfficeIMO.Excel;

using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader("Events.xlsx",
    new ExcelReaderOptions { SheetName = "Events", EnableSinglePassReading = true });

int warnings = 0;
foreach (LogEvent item in reader.RowsAs<LogEvent>()) {
    if (item.Level == "Warning") warnings++;
}
Console.WriteLine($"{warnings} warnings");

record LogEvent(DateTime Time, string Level, string Message);
```

By default the reader validates the whole worksheet before returning the first row. `EnableSinglePassReading` returns rows while the rest of the sheet is still being read; an invalid cell then fails after the rows before it were returned, so use it where a part-way failure can be discarded. It applies on .NET 8 and later to XLSX and XLSB worksheets of 1 MB or more.

### PowerShell

`Import-OfficeExcel -As DataReader` returns one streaming data reader for the sheet. `-ChunkRows` sets how many worksheet rows each streaming chunk requests:

```powershell
$reader = Import-OfficeExcel -Path .\Events.xlsx -Sheet Events -As DataReader -ChunkRows 10000
try {
    $level = $reader.GetOrdinal('Level')
    $warnings = 0
    while ($reader.Read()) {
        if ($reader.GetString($level) -eq 'Warning') { $warnings++ }
    }
    "$warnings warnings"
} finally {
    $reader.Dispose()
}
```

## Untrusted uploads

For files from outside your organization, load with `ExcelLoadOptions.UntrustedDefaults` or set reader limits, so a crafted workbook cannot exhaust memory.
