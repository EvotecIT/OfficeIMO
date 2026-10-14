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

## Project an open workbook

With `OfficeIMO.Reader.Excel` imported, call `document.ReadDocument(sourceName,
options, formatOptions, cancellationToken)` on a caller-owned `ExcelDocument`.
It uses the canonical Reader mapping, snapshots both option sets, including
native read options, and enforces the projection's output budgets. `Full` is the
default. The workbook remains open for subsequent `InspectEmbeddedPayloads` reads.

The source name supplies citations without reopening a file or inventing its
byte hash, length or timestamp. Input and package admission settings cannot
retroactively constrain the earlier native load; configure them when loading
the workbook. Lean projection still requires a modern XML source model.

## Text and metadata projection

`ReaderOptions.Projection = ReaderProjection.TextAndMetadata` supports XLSX,
XLSM, XLTX, XLTM and XLAM. It retains table text, saved values, links and requested
formula/comment/visibility evidence while skipping Markdown rendering and image
asset reads before allocation. Formula inspection does not calculate results or
refresh external connections.

Requested package evidence uses bounded inert payload descriptors. Unknown sizes
remain explicit without an unrequested scan. VBA part/signature-relationship
metadata is retained; binary module decoding is marked
`module-payload-unrequested`.

XLSB, binary XLS families and the separate legacy spreadsheet handlers require
`Full`. The lean profile refuses their native binary-to-XML projection, including
mislabeled modern filenames. Ordinary reads keep their existing format support.
Native parser limits remain separate from Reader's aggregate result budgets.

`OpenXmlMaxPartUncompressedBytes` and `OpenXmlMaxTotalUncompressedBytes` apply
native declared part/aggregate expansion admission before model loading, even
when `Inspection` is null. Configured ceilings produce terminal
`ReaderResourceLimitException` with the native failure retained; they are not
retried as password failures. These are package admission bounds, not a hard
limit on process memory or every decoder's actual work.


Selected Lotus 1-2-3, Quattro Pro, Multiplan, Works, SYLK (`.slk`), and DIF (`.dif`) spreadsheet sources are included in `OfficeIMO.Excel` but remain opt-in at the Reader registration boundary:

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

SYLK/DIF project stored values through the same Excel table reader. SYLK expressions and formatting are omitted with warnings; DIF error markers remain literal text with warnings. The [Excel profile coverage](../OfficeIMO.Excel/README.md#profile-coverage) defines the exact grammar and limits. `LegacySpreadsheetImportOptions.TextEncoding` is snapshotted at registration; when absent, the per-read `ReaderOptions.TextEncoding` applies, then Windows-1252. A Unicode BOM takes precedence. `OfficeIMO.Reader.All` includes these formats when `IncludeLegacySpreadsheet` is enabled, which is the preset default.

## Incremental path reads

`reader.EnumerateChunks(path, options)` yields the existing Excel row chunks on demand and releases the
workbook when the enumerator is disposed. It opens the OfficeIMO.Excel workbook and metadata but avoids
building the complete Reader rich graph. Use `ReadDocument` for workbook metadata and assets.
Stream input and document processors retain the materialized contract.

Registration snapshots `ReaderExcelOptions.ReadOptions`, its culture and execution thresholds.
Caller-supplied delegates remain caller-owned and must be safe for concurrent use.

Applicable native Excel read ceilings, including shared-string limits during
loaded-model projection, become terminal `ReaderResourceLimitException` values
with the native owner, limit, effective maximum and inner exception preserved.
Lazy enumeration, batch, folder and nested ZIP reads retain that classification.
Input and metadata read limits do not retroactively admit an already-loaded
model. Dense-range and XML data-reader buffering options retain their native API
scope; they are not whole-operation allocation guarantees for streamed extraction.

## Inspect stored formulas, comments and hidden state

```csharp
OfficeDocumentReadResult workbook = reader.ReadDocument("review.xlsx", new ReaderOptions {
    Inspection = new ReaderInspectionOptions()
});
foreach (OfficeDocumentMetadataEntry entry in workbook.Metadata) {
    if (entry.Category == OfficeDocumentInspectionCategories.Formula)
        Console.WriteLine($"{entry.SourceObjectId}: {entry.Value}");
}
```

Formula records distinguish the stored formula declaration, projected expression,
stored value/type and interpreted cached value. `hasStoredValue` distinguishes an
absent cache from an available saved value; inspection never recalculates formulas.
Cell notes and threaded comments retain available author, reply and resolution
identifiers, including comments attached to otherwise empty cells. `SheetName` and
`A1Range` select formula/comment evidence. Visibility records retain authored hidden
sheet, row and column definitions for the selected sheets.

Package and content-safety inspection use the loaded owner and original-package
capture. Payload and VBA records remain inert. Inspection metadata is available on
rich or structured reads; incremental row chunks do not carry a complete inspection
record collection. [Reader Core](../OfficeIMO.Reader.Core/README.md#inspect-review-and-saved-content)
defines the shared budgets and terminal resource failures.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
- OfficeIMO dependencies: `OfficeIMO.Reader.Core`, `OfficeIMO.Excel`, and `OfficeIMO.Core`.
- License: MIT.
