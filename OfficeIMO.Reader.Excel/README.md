# OfficeIMO.Reader.Excel

Excel workbook ingestion for `OfficeIMO.Reader.Core`. Install this package when a reader needs XLSX, XLSM, XLSB, or XLS support, then compose it explicitly:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddExcelHandler()
    .Build();
```

For CSV import and export, install the separate `OfficeIMO.Excel.Csv` adapter.
Keeping conversion outside the Reader package avoids pulling CSV behavior into
applications that only need generic Excel document extraction.

Selected Lotus 1-2-3, Quattro Pro, Multiplan, and Works spreadsheet sources are opt-in through the owning legacy package:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddExcelHandler()
    .AddLegacySpreadsheetHandler()
    .Build();

OfficeDocumentReadResult workbook = reader.ReadDocument("archive.wk1");
```

Legacy warnings include the detected profile, structured-versus-salvage quality, and feature-level losses. The handler never executes macros or refreshes external links.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
- OfficeIMO dependencies: `OfficeIMO.Reader.Core`, `OfficeIMO.Excel`, `OfficeIMO.Excel.Legacy`, and `OfficeIMO.Core`.
- License: MIT.
