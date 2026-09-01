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

Selected Lotus 1-2-3, Quattro Pro, Multiplan, and Works spreadsheet sources are included in `OfficeIMO.Excel` but remain opt-in at the Reader registration boundary:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddExcelAndLegacyHandlers(new LegacySpreadsheetImportOptions {
        Limits = new OfficeLegacyImportLimits { MaxInputBytes = 64 * 1024 * 1024 }
    })
    .Build();

OfficeDocumentReadResult workbook = reader.ReadDocument("archive.wk1");
```

The combined registration applies the same immutable legacy options to every legacy spreadsheet route. `AddLegacySpreadsheetHandler(...)` remains available when the normal Excel handler is registered separately or is not needed.

Legacy warnings include the detected profile, structured-versus-salvage quality, and feature-level losses. The handler never executes macros or refreshes external links.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
- OfficeIMO dependencies: `OfficeIMO.Reader.Core`, `OfficeIMO.Excel`, and `OfficeIMO.Core`.
- License: MIT.
