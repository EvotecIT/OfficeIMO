# OfficeIMO.Excel.Csv

`OfficeIMO.Excel.Csv` provides bidirectional CSV and Excel conversion without
coupling either core format package to the other. CSV parsing, schema inference,
compression, and writing remain owned by `OfficeIMO.CSV`; worksheet insertion,
range extraction, and workbook saving remain owned by `OfficeIMO.Excel`.

## Install

```powershell
dotnet add package OfficeIMO.Excel.Csv
```

## Import CSV into Excel

```csharp
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;

using var workbook = ExcelDocument.Create();
ExcelCsvImportResult imported = workbook.ImportCsvFile("sales.csv");
workbook.Save("sales.xlsx");
```

Use `ImportCsvText`, `ImportCsv(Stream)`, or `ImportCsv(CsvDocument)` for other
source shapes. `ExcelCsvImportOptions.LoadOptions` and `ReaderOptions` expose the
canonical CSV option types instead of introducing another parsing model.
`ExcelCsvImportResult.Delimiter` reports the delimiter that was actually used,
including one selected by delimiter detection.

## Export Excel data to CSV

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;

using var workbook = ExcelDocument.Load("sales.xlsx");
workbook["Sales"].SaveAsCsv("sales.csv");
```

Worksheets and ranges can also be returned as CSV text or materialized as a
`CsvDocument`. The adapter transports rows through `IDataReader`/`DbDataReader`;
it does not contain another CSV parser or Excel engine.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
- OfficeIMO dependencies: `OfficeIMO.CSV` and `OfficeIMO.Excel`.
- External dependencies: none beyond those format packages.
- License: MIT.
