# Upgrading OfficeIMO

This guide contains version-to-version changes that require application code, package references, or configuration to change. It is not a release history or a second API manual.

- Use [GitHub Releases](https://github.com/EvotecIT/OfficeIMO/releases) for release notes and downloadable artifacts.
- Use the root and package READMEs for the current public API.
- Use support matrices for current coverage and limits.
- Use this guide when an upgrade no longer compiles or changes an existing workflow.

The repository source is on the coordinated `3.1.x` line; the latest NuGet release is `3.0.3`. Keep every OfficeIMO package in one application on the same published compatibility line and perform a clean restore after changing versions.

| Version in the application | Upgrade path |
| --- | --- |
| `3.0.x` | Apply [3.0 to 3.1](#officeimo-30-to-31). |
| `2.x` | Apply [2.x to 3.0](#officeimo-2x-to-30), then 3.0 to 3.1. |
| `1.x` | Apply [1.x to 2.0](#officeimo-1x-to-20), then each later section in order. |

## OfficeIMO 3.0 to 3.1

### Package changes

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `OfficeIMO.OpenDocument.Pdf` | Install only the required `OfficeIMO.OpenDocument.Odt.Pdf`, `OfficeIMO.OpenDocument.Ods.Pdf`, or `OfficeIMO.OpenDocument.Odp.Pdf` adapter. |
| `OfficeIMO.Reader.Tool` and the `officeimo-reader` executable | `OfficeIMO.Tool` and the `officeimo reader` command area. |
| Public `SixLabors.ImageSharp` color/image helper types | First-party `OfficeIMO.Drawing` types such as `OfficeColor`. |

There is no replacement umbrella OpenDocument PDF package. The focused packages keep Word, Excel, and PowerPoint dependencies out of applications that do not use those routes.

### CSV and Excel tabular reads

CSV and Excel retain separate document models and use the same entry-point grammar:

| Intent | CSV | Excel |
| --- | --- | --- |
| Stream from a path or stream | `CsvDocument.OpenDataReader(...)` | `ExcelDocument.OpenDataReader(...)` |
| Read an already-open document | `csv.CreateDataReader(...)` | `workbook.CreateDataReader(...)` |
| Load an editable model | `CsvDocument.Load(...)` | `ExcelDocument.Load(...)` |

Replace the removed public reader roots as follows:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `CsvDocument.CreateDataReader(pathOrStream, ...)` | `CsvDocument.OpenDataReader(pathOrStream, ...)` |
| Public `CsvDataReader` construction or return types | `DbDataReader` returned by `CsvDocument.OpenDataReader(...)` or `csv.CreateDataReader(...)` |
| `CsvDocument.ReadFieldSpans*`, `CsvDocument.ReadRowFieldSpans*`, and public field-span visitor types | `CsvDocument.OpenDataReader(...)` for streaming, or `Load(...)` / `Parse(...)` for a materialized document |
| `ExcelDocumentReader.Open(...)` | `ExcelDocument.OpenDataReader(...)` |
| `ExcelRead.*`, `ExcelDocument.Read().Sheet().Range()`, or `ExcelSheetReader` | `ExcelDocument.OpenDataReader(...)` for streaming, or `ExcelDocument.Load(...)` for editing |
| Concrete Excel reader return types | `DbDataReader` |

When configuring a streaming CSV read with `CsvLoadOptions`, set `Mode = CsvLoadMode.Stream`; the options object otherwise retains its in-memory default. Excel exposes worksheets as ordered `DbDataReader` results through `NextResult()`.

### PDF conversion and import

The 3.1 PDF adapters use destination-shaped names for general conversion and explicit feature names for narrow recovery.

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `PdfSaveOptions` in `OfficeIMO.Word.Pdf` | `WordPdfSaveOptions` |
| `PdfWordReadOptions` | `PdfWordImportOptions` |
| `PdfRtfReadOptions` | `PdfRtfImportOptions` |
| `PdfPowerPointTableImportOptions` | `PdfPowerPointImportOptions` |
| `PdfPowerPointTableImportReport` / `PdfPowerPointTableImportResult` | `PdfPowerPointConversionReport` / `PdfPowerPointConversionResult` |
| `ImportTablesToPowerPointPresentation` | `ToPowerPointPresentation` |
| `SaveTablesAsPowerPoint` | `SaveAsPowerPoint` |

Excel remains explicitly table-shaped because its PDF adapter recovers detected tables rather than arbitrary page content. Keep using `PdfExcelTableImportOptions`, `PdfExcelTableImportReport`, `PdfExcelTableImportResult`, `ImportTablesToExcelDocument`, and `SaveTablesAsExcel`.

PowerPoint behavior broadens in 3.1: an opened `PdfDocument` creates one rendered page per slide by default. Use `PdfPowerPointImportOptions.CreateEditableTables()` when editable table recovery is the intended result. The visual route does not claim that arbitrary PDF text, vectors, groups, clipping, forms, or annotations become editable PowerPoint objects.

The common conversion grammar is:

| Intent | Shape | Example |
| --- | --- | --- |
| Return a destination model | `To{TargetModel}` | `pdf.ToWordDocument()` |
| Return a model plus diagnostics | `To{TargetModel}Result` | `pdf.ToWordDocumentResult()` |
| Return serialized content | `To{Format}` | `word.ToPdf()` |
| Write a converted artifact | `SaveAs{Format}` | `pdf.SaveAsPowerPoint(...)` |
| Recover a narrow feature | Name the feature | `pdf.SaveTablesAsExcel(...)` |
| Configure forward PDF output | `{Source}PdfSaveOptions` | `WordPdfSaveOptions` |
| Configure reconstruction from PDF | `Pdf{Target}ImportOptions` | `PdfWordImportOptions` |

### Names and behavior

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `AddWorkSheet`, `RemoveWorkSheet`, `CopyWorkSheet`, `ReorderWorkSheet` | `AddWorksheet`, `RemoveWorksheet`, `CopyWorksheet`, `ReorderWorksheet` |
| `MergeWorkSheets`, `JoinWorkSheets`, `CompareWorkSheets` | `MergeWorksheets`, `CompareWorksheets` |
| `RtfDocument.ToHtmlMemoryStream()` | `RtfDocument.ToHtmlStream()` |
| `WordHelpers.ConvertDotXtoDocX(...)` | `WordHelpers.ConvertDotxToDocx(...)` |
| Format-specific color/image helper values exposed as `SixLabors.ImageSharp` types | `OfficeIMO.Drawing` values |

Review these behavioral changes during the upgrade:

- PDF adapters use `PdfResourcePolicy.CreateDefault()` for balanced fidelity. Use `CreatePortableDeterministic()` when conversion must not inspect installed fonts or external resources.
- OpenDocument, AsciiDoc, LaTeX, and semantic OneNote PDF routes return their source-stage mappings separately from PDF layout warnings. Use `HasLoss` or `RequireNoLoss()` for the end-to-end gate.
- Excel and CSV numeric, date, formula-cache, and typed writer values use invariant round-trip formatting on every platform.
- Backend-specific CSV and Excel parser types are no longer public extension points.

## OfficeIMO 2.x to 3.0

### PDF table recovery

OfficeIMO 3.0 renamed its table-only PDF routes so they did not imply full-page reconstruction:

| OfficeIMO 2.x | OfficeIMO 3.0 |
| --- | --- |
| `SaveAsExcel` / `SaveAsExcelAsync` | `SaveTablesAsExcel` / `SaveTablesAsExcelAsync` |
| `ToExcelDocument` / `ToExcelDocumentResult` | `ImportTablesToExcelDocument` / `ImportTablesToExcelDocumentResult` |
| `PdfExcelConversionReport` / `PdfExcelConversionResult` | `PdfExcelTableImportReport` / `PdfExcelTableImportResult` |
| `SaveAsPowerPoint` / `SaveAsPowerPointAsync` | `SaveTablesAsPowerPoint` / `SaveTablesAsPowerPointAsync` |
| `ToPowerPointPresentation` / `ToPowerPointPresentationResult` | `ImportTablesToPowerPointPresentation` / `ImportTablesToPowerPointPresentationResult` |
| `PdfPowerPointConversionReport` / `PdfPowerPointConversionResult` | `PdfPowerPointTableImportReport` / `PdfPowerPointTableImportResult` |

The PowerPoint names broaden again in 3.1 because the default route changes from table-only recovery to one visual slide per PDF page. Apply the 3.0-to-3.1 mappings after completing this section.

### Word, Excel, and EPUB changes

| OfficeIMO 2.x | OfficeIMO 3.0 |
| --- | --- |
| `FormattingHelper.GetFormattedRuns(paragraph)` | `paragraph.GetFormattedRuns()` returning `WordFormattedRun` values |
| `WordListLevel._level` | `WordListLevel.OpenXmlElement` |
| `new WordHelpers()` | Remove the instance; supported `WordHelpers` members are static |
| `WordHelpers.GetNextSdtId(...)` | Remove the call; content-control APIs allocate IDs |
| `InlineRunHelper.AddInlineRuns(...)` | Use the owning converter or explicit paragraph APIs |
| `LegacyXlsLoadResult.Workbook` | `LegacyXlsLoadResult.AdvancedWorkbook` |
| `LegacyXlsLoadResult.ImportReport` or `CreateAdvancedImportReport()` | `LegacyXlsLoadResult.CreateImportReport()` |
| `OfficeIMO.Epub.Html` | `OfficeIMO.Epub.Image` |

The `OfficeIMO.Drawing` target-framework compatibility type `System.Runtime.CompilerServices.IsExternalInit` is internal in 3.0. Remove any application reference to that shim; normal record and `init` usage remains supported.

## OfficeIMO 1.x to 2.0

OfficeIMO 2.0 established the shared lifecycle and result vocabulary used by the current packages.

### Document lifecycle

| Intent | Current API |
| --- | --- |
| Save to an associated destination | `Save()` / `SaveAsync()` |
| Save and associate a path or stream | `Save(pathOrStream)` / `SaveAsync(pathOrStream)` |
| Write a copy without changing the destination | `SaveCopy(...)` / `SaveCopyAsync(...)` |
| Produce bytes | `ToBytes()` |
| Produce a new stream positioned at zero | `ToStream()` |
| Return another format | `To{Format}()` / `To{Format}Result()` |
| Write another format | `SaveAs{Format}()` / `SaveAs{Format}Async()` |

Caller-owned streams remain open. Seekable inputs are read from the beginning and restored to their original position; non-seekable inputs are read forward from their current position. A retained mutable destination must be writable and seekable.

`Async` now identifies real asynchronous I/O or resource resolution. Use synchronous methods for pure parsing, model projection, byte generation, and in-memory formatting. Removed fake-async wrappers should not be recreated in application compatibility layers.

Reusable options contain configuration only. Read diagnostics from the operation result:

- `Value` contains the converted model or encoded output.
- `Report` contains diagnostics and fidelity evidence.
- `HasLoss` reports simplification or omission.
- `RequireValue()` and `RequireNoLoss()` provide fail-fast gates.

### Common member replacements

| Removed member | Replacement |
| --- | --- |
| `WordImage.SaveToFile(...)` | `WordImage.Save(...)` |
| `WordImage.GetBytes()` / `GetStream()` | `ToBytes()` / `OpenRead()` |
| `WordDocument.GetImages()` / `GetImageStreams()` | `GetImageBytes()` / `OpenImageStreams()` |
| `ExcelImage.GetBytes()` | `ExcelImage.ToBytes()` |
| `WordComment.Delete()` | `WordComment.Remove()` |
| `WordTable.AutoFit` | `WordTable.LayoutMode` |
| `ExcelDocument.CreateTableOfContents(...)` | `AddTableOfContents(...)` |
| `ExcelSheet.SetCellValues(...)` | `CellValues(...)` |
| `ExcelSheet.CellValuesParallel(...)` | `CellValues(..., ExecutionMode.Parallel)` |
| `MarkdownDoc.SaveHtml(...)` | `SaveAsHtml(...)` |
| `ToPdfResult()` | `ToPdfDocumentResult()` |
| `HtmlPdfSaveOptions.DocumentOptions` | `HtmlPdfSaveOptions.PdfOptions` |
| PDF `ToWordResult()` | `ToWordDocumentResult()` |
| `PdfSaveResult.ConversionWarnings` | `Warnings` and `Report` |
| `RtfDocument.ToMemoryStream()` | `ToStream()` |
| `ToRtfMemoryStream()` | `ToRtfStream()` |
| `EmailDocument.WriteToBytes()` | `EmailDocument.ToBytes()` |

Image export uses `OfficeImageExportResult` and `OfficeImageExportFormat` from `OfficeIMO.Drawing`. Use `AtDpi(...)` for physical output density, `ForPrint(...)` for the print profile, and an explicit file-conflict policy when replacement or unique naming is required.

PDF adapters use `PdfResourcePolicy` instead of package-specific trust switches. Profiles configure output behavior but do not grant local-file, remote-resource, or host-font access.

## Upgrade checklist

- Upgrade every OfficeIMO package in the application together.
- Remove compatibility wrappers for deleted aliases and compile against the canonical API.
- Replace option-owned diagnostics with operation results.
- Use `ToBytes` / `ToStream` for memory output and `Save` / `SaveAs{Format}` for destinations.
- Keep pure conversion synchronous; await actual file, stream, or remote-resource I/O.
- Review `HasLoss`, omitted-content, and resource-policy diagnostics before accepting converted output.
- Clean package caches, lock files, `bin`, and `obj` outputs when old and new assemblies were restored together.
- Run the application test suite on every supported operating system after the coordinated package upgrade.
