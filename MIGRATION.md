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

OpenDocument forward PDF results expose the typed OpenDocument projection report in `SourceConversionReports` and PDF-layout warnings in `Report`; `ConversionReports` presents both stages in order. Read OpenDocument mappings from `OdfConversionReport.Mappings` instead of looking for the removed synthetic `ODF_*` PDF warnings. Use `HasLoss` or `RequireNoLoss()` for the end-to-end fidelity gate. AsciiDoc, LaTeX, and semantic OneNote PDF routes use the same ordered source-stage and PDF-stage model.

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
| `CsvDocument.ReadFieldSpans*`, `CsvDocument.ReadRowFieldSpans*`, `CsvFieldSpanAction`, `ICsvFieldSpanVisitor`, `ICsvProjectedFieldSpanVisitor`, and `ICsvRowFieldSpanVisitor` | `CsvDocument.OpenDataReader(...)` for streaming, or `Load(...)` / `Parse(...)` for a materialized document |
| `ExcelDocumentReader.Open(...)` | `ExcelDocument.OpenDataReader(...)` |
| `ExcelRead.*`, `ExcelDocument.Read().Sheet().Range()`, or `ExcelSheetReader` | `ExcelDocument.OpenDataReader(...)` for streaming, or `ExcelDocument.Load(...)` for editing |
| Concrete Excel reader return types | `DbDataReader` |

When configuring a streaming CSV read with `CsvLoadOptions`, set `Mode = CsvLoadMode.Stream`; the options object otherwise retains its in-memory default. Excel exposes worksheets as ordered `DbDataReader` results through `NextResult()`.

CSV reader configuration remains in `CsvDataReaderOptions`. Excel reader safety limits remain in `ExcelReadOptions`: `MaxXlsbCells` limits aggregate workbook cells and `MaxDataReaderBufferedCells` limits a reader operation's buffer. Raise either limit only for trusted, intentionally larger workbooks.

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

Use `PdfWordImportOptions.CreateTablesOnly()` for narrow Word table recovery. PowerPoint table details remain available through `PdfPowerPointConversionReport.TableEntries` when the editable-table profile is selected.

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

### PowerPoint lifecycle, composition, and inspection

PowerPoint 3.1 uses the concrete `PowerPointPresentation`, `PowerPointSlide`, and shape types as the editing model. Semantic deck plans, template helpers, and format adapters remain optional workflows over that model.

Replace lifecycle calls as follows:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `Open(path)` | `Load(path)` |
| `OpenRead(path)` | `Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })` |
| `Open(stream, readOnly: true, autoSave: false)` | `Load(stream, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })` |
| `Create(stream, autoSave: false)` | `Create(stream)` |
| Implicit save on dispose | Set `PersistenceMode = DocumentPersistenceMode.SaveOnDispose`, or call `Save()` explicitly |

The designer and deck-composer entry points now route through one plan and one composition call:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `presentation.UseDesigner(...).AddSlides(plan)` | `presentation.Compose(plan, PowerPointCompositionOptions.FromBrief(brief))` |
| `presentation.AddDesignerProcessSlide(...)` | `plan.AddProcess(...)`, then `presentation.Compose(...)` |
| `deck.AddSlidesWithContinuation(plan)` | Set `options.ExpandContinuations = true` (the default), then call `Compose(...)` |
| `deck.AddSlidesWithReport(plan)` | Read the `PowerPointCompositionResult` returned by `presentation.Compose(...)` |

Template ownership and inspection names are also explicit:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `PowerPointPresentation.InspectTemplate(path)` | `PowerPointTemplate.Inspect(path)` |
| `PowerPointPresentation.CreateFromTemplate(...)` | `PowerPointTemplate.CreatePresentation(...)` |
| `presentation.UseTemplateDesigner(...)` | Set `PowerPointCompositionOptions.TemplateLayouts`, then call `Compose(...)` |
| `Preflight()` | `InspectPreflight()` |
| `CreateVisualProofReport()` | `InspectVisuals()` |
| `SaveWithPreflight()` | Call `InspectPreflight()`, apply the required gate, then call `Save()` |

Replace `PowerPointChartData`, `PowerPointScatterChartData`, their series types, and chart-family-specific add methods with shared `OfficeIMO.Drawing.OfficeChartData` plus `AddChart`, `AddChartCm`, `AddChartInches`, or `AddChartPoints`.

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
- Word-to-HTML emits detected run colors and highlights by default. Set `IncludeRunColorStyles` or `IncludeRunHighlightStyles` to `false` only for deliberately style-reduced HTML.
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
| `ImageShapeStyleHelper` | Use the owning image-shape APIs |
| `HorizontalAlignmentHelper` | Use the public alignment properties on the owning paragraph, table, cell, or image API |
| `LegacyXlsLoadResult.Workbook` | `LegacyXlsLoadResult.AdvancedWorkbook` |
| `LegacyXlsLoadResult.ImportReport` or `CreateAdvancedImportReport()` | `LegacyXlsLoadResult.CreateImportReport()` |
| `OfficeIMO.Epub.Html` | `OfficeIMO.Epub.Image` |

Detailed `LegacyXlsImportReport` record-family counters such as `CommentsByObjectType` and `DataValidationsByType` are internal diagnostics rather than public application contracts. Use the stable summary counts, `HasImportErrors`, `HasUnsupportedFeatures`, and the public `Diagnostics`, `UnsupportedFeatures`, `PreservedFeatures`, `UnsupportedSheets`, and `CompoundFeatures` collections. Exhaustive parser telemetry is not exposed as public API.

The `OfficeIMO.Drawing` target-framework compatibility type `System.Runtime.CompilerServices.IsExternalInit` is internal in 3.0. Remove any application reference to that shim; normal record and `init` usage remains supported.

### Legacy DOC and XLS API changes

Legacy Word callers use the normal `WordDocument` lifecycle and explicit conversion policies:

| OfficeIMO 2.x | OfficeIMO 3.0 |
| --- | --- |
| `SaveAs(pathOrStream)` | `SaveCopy(pathOrStream)` |
| `SaveAsByteArray()` | `ToBytes()` |
| `SaveAsMemoryStream()` | `ToStream()` |
| `WasLoadedFromLegacyDoc` | `SourceFormat == WordFileFormat.Doc` |
| `MaxWordDocumentStreamBytes` | `MaxInputBytes` |
| `ReportUnsupportedFeatures` | `ReportUnsupportedContent` |
| positional overwrite conversion flag | `FileConflictPolicy` |
| save-triggered application launch | Call `OpenInApplication(path)` explicitly after a successful save |
| lossy conversion Boolean | `LossPolicy` |

Legacy Excel callers use the same explicit format and policy vocabulary:

| OfficeIMO 2.x | OfficeIMO 3.0 |
| --- | --- |
| `WasLoadedFromLegacyXls` | `SourceFormat == ExcelFileFormat.Xls` |
| `MaxWorkbookStreamBytes` | `MaxInputBytes` |
| `ReportUnsupportedRecords` | `ReportUnsupportedContent` |
| overwrite conversion Boolean | `FileConflictPolicy` |
| save-triggered application launch | Call `OpenInApplication(path)` explicitly after a successful save |
| lossy conversion/save Boolean | `LossPolicy` |
| implicit stream format option | `Save(stream, ExcelFileFormat, options)` or `ToXlsx()` / `ToXls()` |

## OfficeIMO 1.x to 2.0

OfficeIMO 2.0 established the shared lifecycle and result vocabulary used by the current packages.

### Document lifecycle

| Intent | Current API |
| --- | --- |
| Save to an associated destination | `Save()` / `SaveAsync()` |
| Save and associate a path | `Save(path)` / `SaveAsync(path)` |
| Write once to a caller-owned stream | `Save(stream)` / `SaveAsync(stream)` |
| Write a copy without changing the destination | `SaveCopy(...)` / `SaveCopyAsync(...)` |
| Produce bytes | `ToBytes()` |
| Produce a new stream positioned at zero | `ToStream()` |
| Return another format | `To{Format}()` / `To{Format}Result()` |
| Write another format | `SaveAs{Format}()` / `SaveAs{Format}Async()` |

Saving to a caller-owned stream is a one-time write and does not replace the document's associated path or source stream. A later parameterless `Save()` therefore uses the existing association, or throws when the document has none. Caller-owned streams remain open. Seekable inputs are read from the beginning and restored to their original position; non-seekable inputs are read forward from their current position. A retained mutable destination must be writable and seekable.

`Async` now identifies real asynchronous I/O or resource resolution. Use synchronous methods for pure parsing, model projection, byte generation, and in-memory formatting. Removed fake-async wrappers should not be recreated in application compatibility layers.

Removed fake-async methods include in-memory Markdown, HTML, and RTF conversions, byte-returning conversion wrappers, `RtfDocument.ReadAsync(string)`, and `RtfDocument.LoadAsync(byte[])`. Use the synchronous conversion, or use `LoadAsync`, `SaveAsync`, and `SaveAs{Format}Async` only when the source, destination, or resource resolution performs real asynchronous I/O.

Reusable options contain configuration only. Read diagnostics from the operation result:

- `Value` contains the converted model or encoded output.
- `Report` contains diagnostics and fidelity evidence.
- `HasLoss` reports simplification or omission.
- `RequireValue()` and `RequireNoLoss()` provide fail-fast gates.

OpenDocument save methods now return `OdfSaveResult` directly. Replace the discarded-result aliases as follows:

| Removed member | Replacement |
| --- | --- |
| `SaveResult` / `SaveResultAsync` | `Save` / `SaveAsync` returning `OdfSaveResult` |
| `ToBytesResult` | `Serialize` returning `OdfSaveResult` |
| `SaveFlatXmlResult` | `SaveFlatXml` returning `OdfSaveResult` |

Reusable conversion options no longer retain operation state in members such as `LastSaveReport`, `LastSaveDiagnostics`, `ConversionReport`, or `Warnings`. Read that evidence from the returned result.

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
| `OfficeDocumentReadResultSchema.Version` | `OfficeDocumentReadResultSchema.CurrentVersion` |
| `SheetComposer.DefinitionList(...)` | `SheetComposer.PropertiesGrid(...)` |
| `PowerPointUnits.Cm/Mm/Inches/Points(...)` | `FromCentimeters/FromMillimeters/FromInches/FromPoints(...)` |
| `VisioDocument.UseMastersFromTemplate(...)` | `LearnMastersFromVsdx(...)` |
| `OrderedListBlock.ListItems` / `UnorderedListBlock.ListItems` | `Items` |
| `ListItem.Children` | `NestedBlocks` |
| `QuoteBlock.Children` / `DetailsBlock.Children` | `ChildBlocks` |
| `TableCell.Blocks` / `DefinitionListDefinition.Blocks` | `ChildBlocks` |
| `FootnoteDefinitionBlock.Blocks` | `ChildBlocks` |
| tuple-based `DefinitionListBlock.Items` | typed `Groups`, `Entries`, and `AddEntry(...)` |
| `MarkdownDoc.SaveHtml(...)` | `SaveAsHtml(...)` |
| Markdown `VisualTheme` | `Theme` with a shared `MarkdownVisualTheme` |
| `ApplyWordLikeTheme()` | `ApplyDefaultTheme()` |
| `UseFrontMatterVisualTheme` | `UseFrontMatterTheme` |
| `OutlookContact.Email1Address` | `OutlookContact.Email1.Address` |
| phone compatibility properties | `OutlookContact.Phones` |
| `TrackComments` | No replacement; use `TrackChanges` or `Settings.TrackRevisions` for revision tracking. |
| `ToPdfResult()` | `ToPdfDocumentResult()` |
| `HtmlPdfSaveOptions.DocumentOptions` | `HtmlPdfSaveOptions.PdfOptions` |
| `AsciiDocPdfSaveOptions.PdfOptions` | `AsciiDocPdfSaveOptions.MarkdownOptions` |
| `LatexPdfSaveOptions.PdfOptions` | `LatexPdfSaveOptions.MarkdownOptions` |
| `OneNotePdfSaveOptions.PdfOptions` | `OneNotePdfSaveOptions.MarkdownOptions` |
| PDF `ToWordResult()` | `ToWordDocumentResult()` |
| `PdfSaveResult.ConversionWarnings` | `Warnings` and `Report` |
| `RtfDocument.ToMemoryStream()` | `ToStream()` |
| `ToRtfMemoryStream()` | `ToRtfStream()` |
| `SavePdfAsWord()` / `SavePdfAsRtf()` | `SaveAsWord()` / `SaveAsRtf()` on `PdfDocument` |
| `SavePdfTablesAsExcel/Word/PowerPoint()` | `SaveAsExcel()` / `SaveAsWordDocument()` / `SaveAsPowerPoint()` |
| `ToPngResult` / `ToSvgResult` and plural result aliases | `ExportImage()` / `ExportImages()` returning `OfficeImageExportResult` values |
| `PdfImageExportOptions.MaxPages` | `MaximumOutputCount` or `ToImages().WithMaximumPages(...)` |
| `EmailDocument.WriteToBytes()` | `EmailDocument.ToBytes()` |

Format-spelling aliases such as `SaveToPdf`, `SaveAsBytesToPdf`, and generic `WriteToBytes` were removed. Use `SaveAsPdf(...)` for a destination and `ToPdf()` or `ToBytes()` for an in-memory result. Ambiguous `SaveImage` / `SaveAsImage` names were replaced by explicit encodings such as `SaveAsPng(...)`, or by `SaveAsImages(...)` for multi-page and multi-sheet output.

Image export uses `OfficeImageExportResult` and `OfficeImageExportFormat` from `OfficeIMO.Drawing`. Replace the removed scale presets as follows:

| Removed member | Replacement |
| --- | --- |
| `WithDpi(...)` | `AtDpi(...)` for physical output density |
| `ForHighResolution(...)` | `ForPrint(...)` for the print profile |

Image file saves now default to `OfficeImageExportFileConflictPolicy.FailIfExists`. A repeated write to the same path therefore throws unless the caller explicitly selects `Replace` or `CreateUnique` with `OnFileConflict(...)`.

Raster exports share a 50-million-output-pixel default. The default overflow policy reduces scale before allocating the pixel buffer and reports `IMAGE_RASTER_SCALE_REDUCED`; set `RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw` to receive an `OfficeImageExportLimitException` instead.

Image decode and font-fallback diagnostics moved to the shared Drawing result:

| Removed diagnostic | Replacement |
| --- | --- |
| `ExcelImageRasterFormatUnsupported`, `ExcelImageSvgFormatUnsupported`, `ExcelImagePngDecodeUnavailable`, `ExcelHeaderFooterImageUnsupported` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `unsupported-word-image-raster` / `unsupported-word-image-svg` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `unsupported-powerpoint-image-raster` / `unsupported-powerpoint-image-svg` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `HtmlRenderRasterDecoderUnavailable` | `IMAGE_SOURCE_DECODE_FALLBACK` on the final image export result |
| `ExcelCellFontFamilyFallback`, `ExcelChartFontFamilyFallback`, `ExcelHeaderFooterFontFamilyFallback` | `IMAGE_FONT_SUBSTITUTED` |

`IMAGE_SOURCE_DECODED_BY_CALLER_CODEC` records that a caller-supplied `ImageCodec` handled the source.

PDF adapters use `PdfResourcePolicy` instead of package-specific trust switches. Replace the removed switches as follows:

| Removed member | Replacement |
| --- | --- |
| `AllowSystemFontEmbedding` | `ResourcePolicy.AllowSystemFontEmbedding` or `PdfResourcePolicy.CreateTrustedHost()` |
| Markdown `IncludeLocalImages` | `IncludeImages` plus `ResourcePolicy.AllowLocalFileAccess` |
| Markdown `IncludeDataUriImages` | `IncludeImages` plus `ResourcePolicy.AllowDataUris` |

Profiles configure output behavior but do not grant local-file, remote-resource, or host-font access.

Word `IncludePageNumbers` and Excel `IncludeSheetHeadings` now default to `false`; set the corresponding option to `true` when synthetic visible page numbers or worksheet headings are required. PowerPoint no longer exposes `UseSharedVisualSnapshot`: full-slide PDF uses the native PDF renderer, while PNG, SVG, HTML review, and thumbnails use the shared visual snapshot. OneNote PDF conversion accepts one `OneNotePdfSaveOptions` object and returns semantic-projection diagnostics through `ToPdfDocumentResult()`.

## Upgrade checklist

- Upgrade every OfficeIMO package in the application together.
- Remove compatibility wrappers for deleted aliases and compile against the canonical API.
- Replace option-owned diagnostics with operation results.
- Use `ToBytes` / `ToStream` for memory output and `Save` / `SaveAs{Format}` for destinations.
- Keep pure conversion synchronous; await actual file, stream, or remote-resource I/O.
- Review `HasLoss`, omitted-content, and resource-policy diagnostics before accepting converted output.
- Clean package caches, lock files, `bin`, and `obj` outputs when old and new assemblies were restored together.
- Run the application test suite on every supported operating system after the coordinated package upgrade.
