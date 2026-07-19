# OfficeIMO.Reader.Csv - CSV reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Csv)](https://www.nuget.org/packages/OfficeIMO.Reader.Csv)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Csv?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Csv)

`OfficeIMO.Reader.Csv` provides CSV/TSV ingestion for `OfficeIMO.Reader.Core` and CSV exchange helpers for `OfficeIMO.Excel`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Csv
```

## Configure

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddCsvHandler()
    .Build();
```

## Examples

### Read CSV as table-aware chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddCsvHandler(new CsvReadOptions {
        ChunkRows = 100,
        HeadersInFirstRow = true,
        IncludeMarkdown = true
    })
    .Build();

foreach (var chunk in reader.Read("people.csv", new ReaderOptions {
    MaxInputBytes = 25L * 1024L * 1024L,
    MaxTableRows = 100
})) {
    Console.WriteLine($"{chunk.Id}: {chunk.Location.StartLine}-{chunk.Location.EndLine}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read a stream from an upload

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddCsvHandler()
    .Build();

await using var stream = File.OpenRead("upload.tsv");
var chunks = reader.Read(stream, "upload.tsv", new ReaderOptions {
    MaxChars = 4_000,
    ComputeHashes = true
}).ToList();

foreach (var table in chunks.SelectMany(chunk => chunk.Tables ?? Array.Empty<ReaderTable>())) {
    Console.WriteLine($"{table.Rows.Count} row(s)");
}
```

### Import CSV into Excel

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Reader.Csv;

using var workbook = ExcelDocument.Create("report.xlsx");
var result = workbook.ImportDelimitedFile("sales.csv", new ExcelDelimitedImportOptions {
    SheetName = "Sales",
    TableName = "SalesData"
});
```

Worksheet ranges can also be exchanged directly with `sheet.FromCsv(csv)`, `sheet.ToCsv()`, and `sheet.TableToCsv("SalesData")`.

## Output

- CSV/TSV chunks with table-aware output.
- Path and stream dispatch.
- Deterministic chunk IDs and row-based locations.
- `MaxInputBytes` enforcement through shared `ReaderInputLimits`.
- Excel worksheet and table CSV exchange.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.CSV`; parsing stays in the native CSV package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
