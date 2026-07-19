# XLS and XLSX compatibility

OfficeIMO.Excel provides first-party support for Office Open XML `.xlsx` and the supported Excel 97-2003 BIFF binary `.xls` subset. Microsoft Excel, COM automation, and third-party spreadsheet conversion libraries are not runtime dependencies.

This document is the current capability contract. It replaces the implementation roadmap.

## Normal API

Use the same `ExcelDocument` surface for both formats:

```csharp
using OfficeIMO.Excel;

using ExcelDocument workbook = ExcelDocument.Load("input.xls");
Console.WriteLine(workbook.SourceFormat); // ExcelFileFormat.Xls

workbook.Save("output.xlsx");
workbook.Save("copy.xls", new ExcelSaveOptions {
    LossPolicy = ExcelConversionLossPolicy.Allow
});

byte[] xlsx = workbook.ToBytes();
byte[] xls = workbook.ToXls();
```

For an independent copy, use `SaveCopy`. For streams, select `ExcelFileFormat.Xlsx` or `ExcelFileFormat.Xls` explicitly.

For structured file conversion:

```csharp
ExcelDocumentConversionResult result = ExcelDocument.Convert(
    "input.xls",
    "output.xlsx",
    new ExcelDocumentConversionOptions {
        FileConflictPolicy = ExcelConversionFileConflictPolicy.FailIfExists,
        LossPolicy = ExcelConversionLossPolicy.Block
    });

foreach (ExcelConversionDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

The defaults are conservative:

- content determines the source format even when an extension is misleading;
- same-format conversion is rejected;
- existing output is preserved unless `Replace` is selected, and a read-only destination is never replaced;
- known conversion loss blocks conversion and normal saves;
- output is staged and atomically committed;
- cross-family OLE input, such as DOC passed to Excel, is rejected clearly.

Set `LossPolicy = ExcelConversionLossPolicy.Allow` only for reviewed, intentional loss. The same policy exists on conversion and save options.

## XLS import capability

The BIFF reader projects supported content into the normal OfficeIMO workbook model. Current covered families include:

| Family | XLS to XLSX behavior |
|---|---|
| Worksheets, cells, types, cached values, and rich text | Projected |
| Common formulas and shared/array formula structures | Projected or diagnosed when unsupported |
| Fonts, number formats, fills, borders, alignment, and protection | Projected |
| Defined names and external-reference metadata | Projected where representable |
| Comments, hyperlinks, merges, row/column layout, and panes/views | Projected |
| AutoFilter, sort, validation, and conditional formatting | Projected for supported records |
| Page setup, headers/footers, and supported header/footer images | Projected |
| Worksheet images and supported drawings | Projected |
| Chart sheets and their supported chart metadata | Projected as chart-sheet parts |
| Supported table/list definitions and styles | Projected |
| Workbook and document properties | Projected |
| Macro sheets, dialog sheets, VBA modules, and unsupported sheet kinds | Diagnosed and treated as conversion loss |
| VBA projects, embedded OLE objects, and legacy signatures | Diagnosed as preserve-only/loss |
| Unsupported BIFF records or damaged structures | Diagnosed or rejected before output |

Projected chart sheets are not classified as conversion loss. Preserve-only records are exposed separately from projected chart sheets and unsupported sheets.

## Native XLS write capability

The BIFF8 writer handles the tested subset of worksheets, values, supported formulas, common styles and formatting, merges, comments, hyperlinks, names, filters, validations, conditional formatting, workbook/sheet layout and views, protection metadata, table-style metadata, and document properties.

XLSX supports more features and larger limits than BIFF8. Before native XLS output, OfficeIMO checks workbook structure, grid and payload limits, formulas, charts/drawings/images, tables and pivots, connections/query tables, threaded comments, unsupported VML, macro/signature content, and other unsupported destination features. A failed preflight leaves existing output unchanged.

This is practical feature parity, not a claim that every XLSX feature can be encoded in BIFF8.

## Detailed import assessment

```csharp
using OfficeIMO.Excel.LegacyXls;

using LegacyXlsLoadResult load = ExcelDocument.LoadLegacyXlsWithReport("input.xls");
LegacyXlsImportSummary summary = load.Summary;

Console.WriteLine($"Worksheets: {summary.WorksheetCount}");
Console.WriteLine($"Chart sheets: {summary.ChartSheetCount}");
Console.WriteLine($"Preserve-only records: {summary.PreservedFeatureCount}");

load.EnsureNoConversionLoss();
```

Use `load.CreateImportReport()` for compact totals or a Markdown diagnostic report, and `load.AdvancedWorkbook` when an application genuinely needs the neutral BIFF model. Corpus-only aggregation details remain internal so they do not become hundreds of compatibility commitments. Import options consistently use `MaxInputBytes` and `ReportUnsupportedContent`; `Password` is additionally available for supported password-to-open XLS inputs. File conversion always enables unsupported-content discovery—even when a supplied import option disables reporting—because `LossPolicy.Block` must never be bypassed silently. Import options are selected from detected physical content, so limits and passwords still apply when a legacy workbook has a misleading extension.

## Breaking API cleanup

| Removed API | Use |
|---|---|
| `WasLoadedFromLegacyXls` | `SourceFormat == ExcelFileFormat.Xls` |
| `MaxWorkbookStreamBytes` | `MaxInputBytes` |
| `ReportUnsupportedRecords` | `ReportUnsupportedContent` |
| overwrite conversion Boolean | `FileConflictPolicy` |
| save-triggered application launch | Call `OpenInApplication(path)` explicitly after a successful save |
| lossy conversion/save Boolean | `LossPolicy` |
| implicit stream format option | `Save(stream, ExcelFileFormat, options)` or `ToXlsx/ToXls` |

## Validation

The dependency-free automated lane covers normal path/stream load, explicit report load, native write/readback, conversion policy, atomic output, and preflight behavior. Optional desktop Excel validation is skipped unless `OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION` is enabled. When enabled, missing Windows, Excel, or required corpus inputs fail rather than silently passing.
