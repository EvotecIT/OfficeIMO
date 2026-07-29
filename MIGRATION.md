# OfficeIMO migration guide

OfficeIMO releases its supported package family on coordinated compatibility lines. Upgrade every OfficeIMO package in an application together, then perform a clean restore so lock files and cached transitive packages do not retain assemblies from the previous line.

Choose the section for the version currently used by the application:

| Current version | Upgrade path |
| --- | --- |
| `3.0` | Complete [Migrating from OfficeIMO 3.0 to 3.1](#migrating-from-officeimo-30-to-31). |
| `2.x` | Complete the [3.0](#migrating-from-officeimo-2x-to-30) section, then the [3.1](#migrating-from-officeimo-30-to-31) section. |
| `1.x` | Complete the [2.0](#migrating-from-officeimo-1x-to-20), [3.0](#migrating-from-officeimo-2x-to-30), and [3.1](#migrating-from-officeimo-30-to-31) sections in order. |

Do not mix package compatibility lines in one dependency graph. Adapter packages reference their owning document and renderer packages from the same coordinated line.

## Migrating from OfficeIMO 3.0 to 3.1

OfficeIMO 3.1 gives each optional PDF adapter one discoverable surface in both directions. Open a PDF once with `PdfDocument.Open(...)`, then call the destination-shaped method supplied by the package you installed.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

PdfDocument pdf = PdfDocument.Open("source.pdf");
PdfWordConversionResult result = pdf.ToWordDocumentResult();

using OfficeIMO.Word.WordDocument word = result.Value;
word.Save("source.docx");
```

General reverse conversions use destination-shaped names. Excel remains explicit because its
reverse adapter currently recovers detected tables rather than arbitrary page content:

```csharp
pdf.SaveAsPowerPoint("pages.pptx");
pdf.SaveAsHtml("review.html");
pdf.SaveAsRtf("editable.rtf");
pdf.SaveTablesAsExcel("tables.xlsx");
```

Every editable destination also accepts an already loaded `PdfLogicalDocument`. Use that lower-level receiver when you need custom layout analysis or page selection:

```csharp
PdfLogicalDocument selected = pdf.Read.Logical(
    PdfPageSelection.Parse("1-3,5"),
    new PdfTextLayoutOptions { ForceSingleColumn = true });

selected.SaveAsWord("selected.docx");
```

### OpenDocument PDF packages

The 3.0 `OfficeIMO.OpenDocument.Pdf` package pulled Word, Excel, and PowerPoint adapters together. In 3.1 it is replaced by focused packages so applications carry only the route they use:

| Route | Focused package | Reverse entry point |
| --- | --- | --- |
| ODT ⇄ PDF | `OfficeIMO.OpenDocument.Odt.Pdf` | `pdf.ToOdtDocument()` |
| ODS ⇄ PDF | `OfficeIMO.OpenDocument.Ods.Pdf` | `pdf.ToOdsDocument()` |
| ODP ⇄ PDF | `OfficeIMO.OpenDocument.Odp.Pdf` | `pdf.ToOdpPresentation()` |

There is no umbrella or bridge-specific Core package in 3.1. Install the format adapter your application actually uses.

Each reverse result exposes the native PDF import report and the OpenDocument feature-mapping report. Forward ODT/ODS/ODP-to-PDF conversions keep the canonical `PdfDocumentConversionResult`: `SourceConversionReports` contains the typed OpenDocument projection report, `Report` contains PDF-layout warnings, and `ConversionReports` presents both stages in order. `HasLoss` and `RequireNoLoss()` cover every stage, and the same ordered reports flow into `PdfSaveResult` returned by `SaveAsPdf`.

This replaces the 3.0 behavior that flattened OpenDocument feature mappings into synthetic `ODF_*` PDF warnings. Inspect the typed `OdfConversionReport.Mappings` instead; `PdfDocumentConversionResult.Warnings` now describes the PDF stage only.

`HasWarnings` remains the PDF-stage warning flag because source reports have format-specific diagnostic models. Use `HasLoss` for the common end-to-end fidelity gate. Conversion proof can enforce the same rule with `new PdfConversionProofOptions().RequireNoLoss()`.

The same stage-report model applies to AsciiDoc, LaTeX, and semantic OneNote PDF routes: their native Markdown-projection report appears in `SourceConversionReports`, while parser and PDF-layout diagnostics remain in the PDF report. Multi-stage conversions no longer relabel semantic projection findings as if the PDF renderer produced them.

PDF-to-ODS reports non-table page content as loss. PDF-to-ODP defaults to visual pages when the receiver is an opened `PdfDocument`; the lower-level logical receiver supports the editable-table profile because visual rendering needs the original PDF bytes.

### Conversion API grammar

OfficeIMO 3.1 uses one naming grammar across document conversions. The verb describes what the
method returns or where it writes; it does not describe the converter implementation:

| Intent | Canonical shape | Example |
| --- | --- | --- |
| Return the destination model | `To{TargetModel}` | `pdf.ToWordDocument()` |
| Return the model plus diagnostics | `To{TargetModel}Result` | `pdf.ToWordDocumentResult()` |
| Return serialized in-memory content | `To{Format}` | `word.ToPdf()`, `pdf.ToHtml()` |
| Write a converted artifact | `SaveAs{Format}` | `pdf.SaveAsWord(...)`, `pdf.SaveAsPowerPoint(...)` |
| Write asynchronously | `SaveAs{Format}Async` | `pdf.SaveAsRtfAsync(...)` |
| Persist a document in its native format | `Save` / `SaveAsync` | `word.Save(...)` |
| Recover one narrow feature | Name the feature explicitly | `pdf.SaveTablesAsExcel(...)` |
| Configure a forward PDF save | `{Source}PdfSaveOptions` | `WordPdfSaveOptions` |
| Configure the shared PDF writer inside direct save options | `PdfOptions` | `HtmlPdfSaveOptions.PdfOptions` |
| Configure an intermediate conversion stage | `{Intermediate}Options` | `OneNotePdfSaveOptions.MarkdownOptions` |
| Configure semantic reconstruction from PDF | `Pdf{Target}ImportOptions` | `PdfWordImportOptions` |
| Describe a general reverse conversion | `Pdf{Target}ConversionResult` / `Report` | `PdfPowerPointConversionResult` |
| Describe narrow table recovery | `Pdf{Target}TableImportResult` / `Report` | `PdfExcelTableImportResult` |

Image export follows the same result-versus-write distinction while accounting for page
cardinality: `ToImage()` opens the fluent export builder, `ExportImage()` returns one structured
render result, `SaveAsPng` / `SaveAsJpeg` / similar methods write an explicit encoding, and
`SaveAsImages()` writes a multi-page or multi-sheet set. The public surface does not use
`SaveImage` or the ambiguous singular `SaveAsImage`.

Target names use .NET casing (`Pdf`, `Html`, `Rtf`, `Odt`, `Ods`, `Odp`, and `PowerPoint`).
Async suffixes are reserved for methods that perform asynchronous I/O or genuinely asynchronous
resource work.

### PDF bridge API replacements

The 3.1 boundary removes the overlapping 3.0 names instead of retaining aliases.

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `PdfSaveOptions` in `OfficeIMO.Word.Pdf` | `WordPdfSaveOptions` |
| `PdfWordReadOptions` | `PdfWordImportOptions` |
| `PdfRtfReadOptions` | `PdfRtfImportOptions` |
| `PdfPowerPointTableImportOptions` | `PdfPowerPointImportOptions` |
| `PdfPowerPointTableImportReport` / `Result` | `PdfPowerPointConversionReport` / `Result` |
| `ImportTablesToPowerPointPresentation` | `ToPowerPointPresentation` |
| `SaveTablesAsPowerPoint` | `SaveAsPowerPoint` |

These are not symmetrical renames. A forward option such as `WordPdfSaveOptions` configures saving
the Word source as PDF. A reverse option such as `PdfWordImportOptions` configures how the Word
destination is reconstructed from PDF. Result and report names then describe the breadth of the
operation: `Conversion` for a general destination route and `TableImport` for table-only recovery.

Excel's 3.0 table-specific names remain unchanged in 3.1:
`PdfExcelTableImportOptions`, `PdfExcelTableImportReport`, `PdfExcelTableImportResult`,
`ImportTablesToExcelDocument`, and `SaveTablesAsExcel`. Keeping that narrow contract prevents a
future broader `ToExcelDocument` route from either overstating today's behavior or requiring
another breaking rename.

PowerPoint is intentionally different. Its 3.0 table-specific names described a table-only
adapter; the 3.1 default creates a slide for every PDF page, so the broader destination-shaped
name now represents broader behavior rather than reversing a rename for cosmetic consistency.

`PdfWordImportOptions.CreateTablesOnly()` and `PdfPowerPointImportOptions.CreateEditableTables()`
select narrower reconstruction profiles where the general destination facade already represents
more than tables. PowerPoint table details are exposed through
`PdfPowerPointConversionReport.TableEntries`.

### PDF reverse-route output

| Route | Default output | Editable | Important limit |
| --- | --- | --- | --- |
| PDF → Word | Semantic headings, paragraphs, lists, tables, supported images and links | Yes | Not fixed-layout page reconstruction |
| PDF → Excel | Detected tables and structured data | Yes | Non-table page content is reported, not placed on a worksheet canvas |
| PDF → PowerPoint | One rendered PDF page per slide | Slide images are movable; page internals are not editable | Managed renderer capability diagnostics identify simplifications |
| PDF → PowerPoint with `EditableTables` | Detected tables on editable slides | Yes | Other page content is reported as omitted |
| PDF → RTF | Semantic headings, paragraphs, lists, page breaks, and detected run styling | Yes | Tables, images, links, and form widgets currently produce loss diagnostics |
| PDF → HTML | Semantic HTML or positioned review HTML | Yes | Neither profile is a browser clone of an arbitrary PDF renderer |
| PDF → Markdown | Logical readable text through `pdf.Read.Markdown(...)` | Yes | Intended for portable text, not visual fidelity |
| PDF → ODT | PDF → Word → ODT semantic composition | Yes | Inspect both stage reports; fixed page layout is not reconstructed |
| PDF → ODS | PDF tables → Excel → ODS composition | Yes | Non-table content is explicitly reported as omitted |
| PDF → ODP | Visual PDF pages → PowerPoint → ODP composition | Slide images are movable | Editable-table mode remains available; arbitrary PDF internals are not reconstructed |

Word and RTF semantic import consume shared `PdfLogicalTextRun` fragments. Those fragments preserve detected source color, font size, and best-effort bold/italic classification without making each destination adapter realign raw PDF spans independently.

### PowerPoint import modes

The opened-PDF PowerPoint route defaults to visual pages because a page image is a more useful and honest general PDF-to-slide result than returning only detected tables:

```csharp
PdfDocument pdf = PdfDocument.Open("handout.pdf");
PdfPowerPointConversionReport report = pdf.SaveAsPowerPoint("handout.pptx");
```

For editable table recovery:

```csharp
var options = PdfPowerPointImportOptions.CreateEditableTables();
options.MaxRowsPerSlide = 18;
options.MaxColumnsPerSlide = 6;

PdfPowerPointConversionReport report = pdf.SaveAsPowerPoint(
    "handout-tables.pptx",
    options);
```

The visual mode is the foundation for a later hybrid mode with editable text and image layers. Arbitrary PDF vectors, groups, clipping, forms, annotations, and presentation animations are not claimed as editable PowerPoint objects.

### PDF resource defaults

`PdfResourcePolicy.CreateDefault()` is the balanced fidelity default for PDF adapter packages. It permits installed-font and document-font embedding while continuing to deny arbitrary local-file and remote-resource access.

Use `PdfResourcePolicy.CreatePortableDeterministic()` for untrusted or reproducible jobs that must not inspect host fonts. Use `CreateTrustedHost()` only when a conversion intentionally resolves local or remote resources.

Word-to-HTML now emits detected run colors and highlights by default. Set `IncludeRunColorStyles` or `IncludeRunHighlightStyles` to `false` only when a deliberately style-reduced HTML result is required.

### Reverse-conversion roadmap

Reverse routes remain supported and should expand where the destination model can represent useful content:

- PDF → Excel: improve table continuation, repeated-header recognition, typed values, and bounded positioned-cell recovery; do not present arbitrary page art as a workbook.
- PDF → PowerPoint: add a hybrid visual/editable mode, then reconstruct bounded text boxes and supported image layers while retaining the rendered page as an optional reference.
- PDF → Word and RTF: extend shared run reconstruction, table/image coverage, and positioning diagnostics before attempting broad page-layout claims.
- PDF → HTML: keep semantic and positioned profiles explicit, and improve shared asset/style diagnostics rather than merging them into an ambiguous default.

Each expansion needs an artifact test and a truthful report for content that remains simplified or omitted. The [PDF conversion support matrix](Docs/officeimo.pdf-conversion-support-matrix.md) records the current direct, composed, and planned routes.

## Migrating from OfficeIMO 2.x to 3.0

OfficeIMO 3.0 aligns the supported package set on one release line, makes table-only PDF recovery explicit, and removes public access to implementation details that applications should not have needed.

```xml
<PackageReference Include="OfficeIMO.Word" Version="3.0.0" />
<PackageReference Include="OfficeIMO.Excel" Version="3.0.0" />
<PackageReference Include="OfficeIMO.Pdf" Version="3.0.0" />
```

### PDF table imports

The Excel and PowerPoint PDF adapters recover detected logical tables. They do not reproduce every PDF page element. Their 3.0 names state that contract.

| OfficeIMO 2.x | OfficeIMO 3.0 |
| --- | --- |
| `SaveAsExcel` / `SaveAsExcelAsync` | `SaveTablesAsExcel` / `SaveTablesAsExcelAsync` |
| `ToExcelDocument` | `ImportTablesToExcelDocument` |
| `ToExcelDocumentResult` | `ImportTablesToExcelDocumentResult` |
| `PdfExcelConversionReport` | `PdfExcelTableImportReport` |
| `PdfExcelConversionResult` | `PdfExcelTableImportResult` |
| `SaveAsPowerPoint` / `SaveAsPowerPointAsync` | `SaveTablesAsPowerPoint` / `SaveTablesAsPowerPointAsync` |
| `ToPowerPointPresentation` | `ImportTablesToPowerPointPresentation` |
| `ToPowerPointPresentationResult` | `ImportTablesToPowerPointPresentationResult` |
| `PdfPowerPointConversionReport` | `PdfPowerPointTableImportReport` |
| `PdfPowerPointConversionResult` | `PdfPowerPointTableImportResult` |

Load a logical PDF, then call the table-specific adapter:

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

PdfLogicalDocument source = PdfLogicalDocument.Load("report.pdf");
PdfExcelTableImportReport report = source.SaveTablesAsExcel("tables.xlsx");

if (report.HasOmittedPageContent) {
    Console.WriteLine("The source also contains content outside detected tables.");
}
```

`HasLoss` means a detected table was truncated by an import limit. `HasOmittedPageContent` means the source also contains non-table text, vector graphics, images, links, forms, annotations, or page actions that the table-only adapter does not import. Use `SourceScope` for the counts behind that decision. Use Word/RTF semantic conversion or image rendering when a full-page representation is the goal.

### Word public surface

Several helper types were implementation details rather than stable application APIs:

| OfficeIMO 2.x | OfficeIMO 3.0 |
| --- | --- |
| `FormattingHelper.GetFormattedRuns(paragraph)` | `paragraph.GetFormattedRuns()` returning `WordFormattedRun` values |
| `WordListLevel._level` | `WordListLevel.OpenXmlElement` |
| `new WordHelpers()` | Remove the instance; `WordHelpers` is static and its supported methods are called directly |
| `WordHelpers.GetNextSdtId(...)` | Removed; content-control APIs allocate valid IDs internally |
| `InlineRunHelper.AddInlineRuns(...)` | Use the owning converter or explicit paragraph APIs |
| `ImageShapeStyleHelper` | Use the owning image shape APIs |
| `HorizontalAlignmentHelper` | Use the public alignment properties on the owning paragraph, table, cell, or image API |

For Markdown, parse the document through `OfficeIMO.Word.Markdown` instead of using the old inline-run helper:

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using WordDocument document = MarkdownReader.Parse(markdown).ToWordDocument();
```

`ConvertDotxToDocx` now resolves relative template paths before constructing the package URI, so relative and absolute template paths follow the same behavior.

### Legacy XLS import reports

| OfficeIMO 2.x | OfficeIMO 3.0 |
| --- | --- |
| `LegacyXlsLoadResult.Workbook` | `LegacyXlsLoadResult.AdvancedWorkbook` |
| `LegacyXlsLoadResult.ImportReport` | `LegacyXlsLoadResult.CreateImportReport()` |
| `LegacyXlsLoadResult.CreateAdvancedImportReport()` | `LegacyXlsLoadResult.CreateImportReport()` |
| Detailed `LegacyXlsImportReport` record-family counters | Stable summary counts and issue collections |

`AdvancedWorkbook` is the public imported workbook. The low-level `Workbook` projection and exhaustive parser telemetry are internal in 3.0. `CreateImportReport()` returns the cached public report with the stable summary counts and the derived `HasImportErrors` and `HasUnsupportedFeatures` indicators. Detailed record-family counters remain available to OfficeIMO's import implementation and tests without becoming permanent public API.

### EPUB image export package

The EPUB-to-image adapter is now named for its result:

```text
OfficeIMO.Epub.Html  ->  OfficeIMO.Epub.Image
```

Update both the package reference and namespace imports. The adapter still retains EPUB chapter HTML and package resources internally and renders through the shared HTML image pipeline; the rename does not introduce another HTML renderer.

### Compatibility shim visibility

`OfficeIMO.Drawing` no longer exports `System.Runtime.CompilerServices.IsExternalInit` from its `netstandard2.0` and `net472` assets. That type was a compiler compatibility shim, not an OfficeIMO API. OfficeIMO still supplies an internal shim where the target framework needs one, so record and `init` usage in applications is unaffected. Remove any direct reference to the OfficeIMO-provided shim.

### Package and dependency ownership

OfficeIMO 3.0 keeps format ownership in the existing document, renderer, and adapter projects. There is no new catch-all core package. Small adapter packages such as PDF or image exporters remain thin surfaces over the owning parser and renderer, which avoids duplicating conversion logic or forcing unrelated dependencies into document packages.

The [3.0 public API review](Docs/officeimo-3.0-public-api-review.md) records the assembly-level comparison used to confirm the changed coordinated public surfaces.

## Migrating from OfficeIMO 1.x to 2.0

The coordinated `2.0.0` cleanup marked one compatibility boundary: applications upgraded their OfficeIMO package set together instead of mixing `1.x` and `2.x` packages.

This release is a coordinated breaking cleanup across the OfficeIMO solution. It removes compatibility aliases, duplicate infrastructure, misleading async methods, and option-owned operation state. Consumers should migrate to the canonical APIs below instead of recreating removed names in wrappers.

### Package architecture

`OfficeIMO.Drawing` remains the small shared foundation for document packages. It already owns the cross-format types required by Word, Excel, PowerPoint, Visio, HTML, PDF, fonts, colors, images, charts, lifecycle options, stream helpers, and export results. There is no additional `OfficeIMO.Core` package and no `.Drawing` to `.Core` rename in this release.

The ownership rules are:

- native format packages own parsing, loading, editing, validation, and serialization for their format;
- adapter packages project one native model into another and do not implement a second parser or document brain;
- `OfficeIMO.Reader.Core` owns normalized read orchestration and contracts, while format-specific Reader packages register typed handlers;
- `OfficeIMO.Html` owns the canonical HTML source model, resource policy, media filtering, and render scene;
- shared colors, fonts, images, stream contracts, lifecycle options, and image export results live in `OfficeIMO.Drawing`;
- `OfficeIMO.Security` owns neutral CMS/X.509/RFC 3161 operations, while each format package owns only its signed-artifact orchestration.

The former compiled `OfficeIMO.Shared` implementation layer is gone. `OfficeIMO.SharedSource` is source-only, and reusable runtime behavior has an explicit owner.

### Persistence lifecycle

Mutable document packages use one vocabulary:

| Intent | Canonical API |
| --- | --- |
| Save to the associated destination | `Save()` / `SaveAsync()` |
| Save and associate a path or stream | `Save(pathOrStream)` / `SaveAsync(pathOrStream)` |
| Write a copy without changing the associated destination | `SaveCopy(path)` / `SaveCopyAsync(path)` |
| Produce bytes without changing document state | `ToBytes()` |
| Produce a new stream positioned at the beginning | `ToStream()` |
| Export another format | `To{Format}()` or `To{Format}Result()` |
| Write another format | `SaveAs{Format}()` / `SaveAs{Format}Async()` |

There are no format-spelling variants such as `SaveToPdf`, `SaveAsBytesToPdf`, or `WriteToBytes`. `SaveAs{Format}` always writes to a destination. `To{Format}` returns an in-memory value. Result-bearing conversions expose evidence instead of storing it in reusable options.

OpenDocument saves now return their evidence directly:

```csharp
OdfSaveResult saved = document.Save("output.odt");
saved.RequireNoLoss();

OdfSaveResult serialized = document.Serialize();
byte[] bytes = serialized.RequireValue();
```

`OdfSaveResult` exposes `Value`, `Report`, `HasLoss`, `RequireValue()`, and `RequireNoLoss()`. The discarded-result aliases `SaveResult`, `SaveResultAsync`, `ToBytesResult`, and `SaveFlatXmlResult` were removed. `Save`, `SaveAsync`, `SaveCopy`, `SaveFlatXml`, and `Serialize` are the result-bearing APIs.

### Stream ownership

Caller-owned streams are never disposed by OfficeIMO.

- A seekable input is read from the beginning and restored to its original position.
- A non-seekable input is read from its current position to the end.
- A returned stream is new, seekable, and positioned at zero.
- A stream retained as a mutable document destination must be writable and seekable so a later parameterless `Save()` can replace the complete artifact.
- A one-time `Save(stream)` does not silently redirect future parameterless saves unless that document's create/load lifecycle explicitly associates the stream.

These rules are shared across synchronous and asynchronous reads. Cancellation restores a seekable caller stream before the cancellation escapes.

### Async contract

`Async` means the operation performs asynchronous I/O or asynchronous external resource resolution. Pure parsing, model projection, byte generation, and in-memory formatting remain synchronous.

Remote image and stylesheet operations are async-only. For example:

```csharp
HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
HtmlToWordResult converted = await source.ToWordDocumentResultAsync(options, cancellationToken);
```

The synchronous HTML-to-Word API is deliberately offline-only. It accepts embedded and local resources allowed by the operation policy but rejects an import that would perform HTTP I/O.

Removed fake-async methods include in-memory Markdown/HTML/RTF conversions, byte-returning conversion wrappers, `RtfDocument.ReadAsync(string)`, and `RtfDocument.LoadAsync(byte[])`. Use the synchronous conversion, or use `LoadAsync`, `SaveAsync`, and `SaveAs{Format}Async` when the source or destination performs real I/O.

### Conversion results and diagnostics

Reusable option objects contain configuration only. They no longer retain `LastSaveReport`, `LastSaveDiagnostics`, `ConversionReport`, or `Warnings` from a previous operation.

Structured conversion results consistently provide:

- `Value` for the converted model or encoded output;
- `Report` for diagnostics and fidelity evidence;
- `HasLoss` when the conversion simplified or omitted content;
- `RequireValue()` and `RequireNoLoss()` where failing fast is useful.

The canonical forward PDF result method is `ToPdfDocumentResult()`. Reverse PDF adapters extend `PdfDocument` and `PdfLogicalDocument` with destination-shaped methods such as `ToWordDocumentResult()`, `ToPowerPointPresentationResult()`, and `ToRtfDocumentResult()`.

`SaveAsPdf` now returns structured save evidence across Word, Excel, PowerPoint, HTML, Markdown, and RTF PDF adapters. `ToPdf()` remains the direct encoded-byte convenience API. Launching or opening a generated PDF is application behavior and is not part of saving.

RTF bridges use `RtfConversionResult<T>`. PDF save attempts expose their report, warnings, warning state, and write outcome rather than mutating the conversion options.

### Image export

Word, Excel, PowerPoint, Visio, HTML, email, EPUB, OneNote, PDF, and the ODT/ODS/ODP bridges use `OfficeImageExportResult` and `OfficeImageExportFormat` from `OfficeIMO.Drawing`.

```csharp
HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
OfficeImageExportResult png = source.ExportImage(OfficeImageExportFormat.Png, options);
OfficeImageExportResult saved = source.SaveAsPng("preview.png", options);
```

`ToPng()`, `ToJpeg()`, `ToTiff()`, and `ToWebp()` return encoded bytes; `ToSvg()` returns SVG text. `ExportImage()` and `ExportImages()` return encoded output, dimensions, format, density, source metadata, and diagnostics. Format-specific save methods and the fluent `As...().Save(...)` surface write to a path or stream and return the same structured evidence. The redundant `ToPngResult`, `ToSvgResult`, and plural result aliases were removed.

Every result validates that its encoded bytes and dimensions match the declared format and dimensions. `DpiX`, `DpiY`, `PhysicalWidthInches`, `PhysicalHeightInches`, and `EncodedLength` are derived from the encoded payload. PNG, JPEG, and TIFF write density metadata through the shared encoder.

Shared options own `MaximumRasterPixels`, `RasterOverflowBehavior`, `ImageCodec`, `RasterEncoding`, `TargetDpi`, `Fonts`, `Policy`, `Progress`, aggregate batch limits, and maximum concurrency. Document-specific option types inherit and clone those settings instead of redeclaring them. The shared default is 50 million output pixels per raster. The default overflow policy reduces scale before allocating a pixel buffer and emits `IMAGE_RASTER_SCALE_REDUCED`. Set `RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw` to receive an `OfficeImageExportLimitException` with requested and allowed dimensions.

Use `AtDpi(...)` for physical output density and `ForPrint(...)` for the print profile. `WithDpi(...)` and `ForHighResolution(...)` were removed because they used inconsistent scaling rules across packages. `WithScale(...)` remains for callers that intentionally work in renderer-relative scale.

File saves now fail when the destination already exists. Select `Replace` or `CreateUnique` explicitly:

```csharp
OfficeImageExportResult saved = document
    .ToImage()
    .AsPng()
    .OnFileConflict(OfficeImageExportFileConflictPolicy.CreateUnique)
    .Save("preview");

Console.WriteLine(saved.SavedPath);
```

The returned path is absolute and includes any appended canonical extension or unique suffix. Direct result saves use the same `OfficeImageExportFileConflictPolicy`.

Batch builders now support `ExportEach(...)` / `ExportEachAsync(...)`, cancellation, progress, deterministic bounded concurrency, and aggregate limits for count, raster pixels, and encoded bytes. Use `SaveFiles(...)` / `SaveFilesAsync(...)` to return path/metadata/diagnostics without retaining every encoded payload.

Image diagnostics now include `OfficeImageExportLossKind`. `OfficeImageExportPolicy` can reject all loss, omissions, failures, or selected codes before a direct or fluent export is returned or saved. Missing requested fonts use the shared `IMAGE_FONT_SUBSTITUTED` code. Supply intended TrueType faces through `WithFont(...)`, `WithFonts(...)`, or `OfficeImageExportOptions.Fonts`.

Format-neutral SVG image export now uses whole-pixel `px` root dimensions so its encoded dimensions match `OfficeImageExportResult.Width` and `Height`. The lower-level `OfficeDrawingSvgExporter.ToSvg(drawing, scale)` overload retains its point-based legacy surface; choose `OfficeSvgSizeUnit.Point` explicitly when a non-image Drawing workflow needs points.

PDF exposes the same canonical surface:

```csharp
PdfReadDocument loaded = PdfReadDocument.Open(pdfBytes);
loaded.ToImages()
    .Pages("2,1")
    .AtDpi(144)
    .AsWebp()
    .Save("pages");
```

`PdfDocumentConversionResult` is the one paged-image adapter for any source that already converts to the first-party PDF model. It keeps Markdown, AsciiDoc, LaTeX, RTF, OneNote, Word, Excel, PowerPoint, or HTML conversion warnings on every exported page:

```csharp
IReadOnlyList<OfficeImageExportResult> pages = markdown
    .ToPdfDocumentResult()
    .ToImages()
    .AsPng()
    .Export();
```

`PdfImageExportOptions.MaxPages` was removed because it duplicated the Drawing-owned batch budget. Set `MaximumOutputCount` directly or use `ToImages().WithMaximumPages(...)`; both now enforce the same limit before any selected page is rendered.

Use `PdfReadPage.ToDrawing()` when a caller needs the intermediate `OfficeDrawing` scene. The older `PdfPageRenderResult` batch remains a low-level inspection/OCR/verification contract behind the fluent reader facade because it carries per-page elapsed time, continue-on-error state, and typed PDF capability diagnostics; it is not the general five-format export API.

ODT, ODS, and ODP direct image extensions live in their existing Word/Excel/PowerPoint OpenDocument adapter packages and attach ODF conversion diagnostics to every image. `OfficeIMO.Epub.Image` projects retained EPUB chapter HTML/resources through the HTML renderer. The email bridge selects HTML, RTF, or text bodies and resolves allowed inline MIME resources through the same HTML resource pipeline.

### HTML source ownership

Raw HTML is parsed once into `HtmlConversionDocument`. Direct PDF/image rendering and Word, Markdown, RTF, Excel, and PowerPoint adapters consume that native source model.

```csharp
HtmlConversionDocument source = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
    BaseUri = new Uri("https://example.test/reports/"),
    UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
});

byte[] pdf = source.ToPdf(pdfOptions);
OfficeImageExportResult image = source.ExportImage(OfficeImageExportFormat.Png, imageOptions);
MarkdownDoc markdown = source.ToMarkdownDocument(markdownOptions);
```

The source model preserves the caller base URI, document `<base>` semantics, source DOM, policy diagnostics, and profile media intent. Renderers evaluate media queries against their real viewport or page dimensions. Adapter-specific element filters run before that adapter resolves URLs. This prevents duplicate parsers and inconsistent resource decisions.

### Reader ownership

Use an immutable `OfficeDocumentReader` built from explicit format handlers:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddRtfHandler()
    .AddPdfHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument(path, options, cancellationToken);
```

Native format packages own parsing. Reader adapters translate native models into `OfficeDocumentReadResult`; they do not expose parallel public parser classes. Reader options are reusable configuration, and diagnostics are returned by the read operation.

`OfficeDocumentReadResultSchema.CurrentVersion` is the schema constant. The ambiguous `Version` alias was removed.

### Theme ownership

Markdown HTML and PDF use one cross-format `MarkdownVisualTheme` through `Theme`. PDF-only visual settings remain in `MarkdownPdfSaveOptions.PdfTheme`.

```csharp
var htmlOptions = new HtmlOptions { Theme = MarkdownVisualTheme.Report() };
var pdfOptions = new MarkdownPdfSaveOptions { Theme = MarkdownVisualTheme.Report() };
```

The canonical helpers are `ApplyDefaultTheme()` and `UseFrontMatterTheme`. `VisualTheme`, `ApplyWordLikeTheme()`, and `UseFrontMatterVisualTheme` were removed.

Visio separates two different concepts:

- `VisioStyleTheme` describes reusable diagram styling;
- `VisioPackageTheme` represents the theme stored in a Visio package.

Layout settings remain layout options and are not duplicated as themes. Office colors and hexadecimal formatting are owned by `OfficeIMO.Drawing`; Word and Excel no longer carry duplicate color helpers.

### Image export diagnostics

Source-image decode policy now belongs to `OfficeIMO.Drawing` across Word, Excel, PowerPoint, HTML, OneNote, Visio, and PDF image export. Family-specific preflight warnings that claimed an image was skipped have been removed because the final renderer may decode it through Drawing, a caller-supplied `ImageCodec`, or a visible fallback.

Use the shared result diagnostics instead:

| Removed diagnostic | Replacement |
| --- | --- |
| `ExcelImageRasterFormatUnsupported` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `ExcelImageSvgFormatUnsupported` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `ExcelImagePngDecodeUnavailable` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `ExcelHeaderFooterImageUnsupported` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `unsupported-word-image-raster` / `unsupported-word-image-svg` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `unsupported-powerpoint-image-raster` / `unsupported-powerpoint-image-svg` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `HtmlRenderRasterDecoderUnavailable` | `IMAGE_SOURCE_DECODE_FALLBACK` on the final image export result |
| `ExcelCellFontFamilyFallback` | `IMAGE_FONT_SUBSTITUTED` |
| `ExcelChartFontFamilyFallback` | `IMAGE_FONT_SUBSTITUTED` |
| `ExcelHeaderFooterFontFamilyFallback` | `IMAGE_FONT_SUBSTITUTED` |

`IMAGE_SOURCE_DECODED_BY_CALLER_CODEC` is informational proof that `ImageCodec` handled the source. When no codec succeeds, the renderer keeps the content visible with a placeholder or a documented family-specific artwork fallback; it no longer emits a warning that says content was omitted when it was not. Drawing can rasterize its bounded SVG subset directly; unsupported SVG features continue through the caller codec or the diagnosed fallback.

### Canonical member names

| Removed member | Replacement |
| --- | --- |
| `WordImage.SaveToFile(...)` | `WordImage.Save(...)` |
| `WordImage.GetBytes()` / `GetStream()` | `ToBytes()` / `OpenRead()` |
| `WordDocument.GetImages()` / `GetImageStreams()` | `GetImageBytes()` / `OpenImageStreams()` |
| `ExcelImage.GetBytes()` | `ExcelImage.ToBytes()` |
| `WordComment.Delete()` | `WordComment.Remove()` |
| `WordTable.AutoFit` | `WordTable.LayoutMode` |
| `AddWorkSheet`, `RemoveWorkSheet`, `CopyWorkSheet`, `ReorderWorkSheet` | `AddWorksheet`, `RemoveWorksheet`, `CopyWorksheet`, `ReorderWorksheet` |
| `MergeWorkSheets`, `JoinWorkSheets`, `CompareWorkSheets` | `MergeWorksheets`, `CompareWorksheets` |
| `ExcelDocument.CreateTableOfContents(...)` | `AddTableOfContents(...)` |
| `ExcelSheet.SetCellValues(...)` | `CellValues(...)` |
| `ExcelSheet.CellValuesParallel(...)` | `CellValues(..., ExecutionMode.Parallel)` |
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
| `OutlookContact.Email1Address` | `OutlookContact.Email1.Address` |
| phone compatibility properties | `OutlookContact.Phones` |
| `TrackComments` | no replacement; use `TrackChanges` or `Settings.TrackRevisions` for revision tracking |
| `ToPdfResult()` | `ToPdfDocumentResult()` |
| `HtmlPdfSaveOptions.DocumentOptions` | `HtmlPdfSaveOptions.PdfOptions` |
| `AsciiDocPdfSaveOptions.PdfOptions` | `AsciiDocPdfSaveOptions.MarkdownOptions` |
| `LatexPdfSaveOptions.PdfOptions` | `LatexPdfSaveOptions.MarkdownOptions` |
| `OneNotePdfSaveOptions.PdfOptions` | `OneNotePdfSaveOptions.MarkdownOptions` |
| PDF `ToWordResult()` | `ToWordDocumentResult()` |
| `PdfSaveResult.ConversionWarnings` | `Warnings` and `Report` |
| `RtfDocument.ToMemoryStream()` | `ToStream()` |
| `RtfDocument.ToHtmlMemoryStream()` | `ToHtmlStream()` |
| `ToRtfMemoryStream()` | `ToRtfStream()` |
| `SavePdfAsWord()` / `SavePdfAsRtf()` | `SaveAsWord()` / `SaveAsRtf()` on `PdfDocument` |
| `SavePdfTablesAsExcel/Word/PowerPoint()` | `SaveAsExcel()` / `SaveAsWordDocument()` / `SaveAsPowerPoint()` |
| `WordHelpers.ConvertDotXtoDocX(...)` | `ConvertDotxToDocx(...)` |
| `EmailDocument.WriteToBytes()` | `EmailDocument.ToBytes()` |

Generic file-copy, file-lock probing, duplicate color helpers, public internal save writers, and other APIs with no useful Office document contract were removed rather than renamed.

### PDF converter trust and fidelity defaults

All PDF adapter options now use `PdfResourcePolicy`. The balanced default enables installed fonts and bounded data URI/package resources for document fidelity while denying arbitrary local-file and remote resolver access. For fully reproducible or untrusted conversion, set `PdfResourcePolicy.CreatePortableDeterministic()`. For trusted inputs that intentionally use local or remote resources, set:

```csharp
options.ResourcePolicy = PdfResourcePolicy.CreateTrustedHost();
```

The following duplicate trust switches were removed:

| Removed member | Replacement |
| --- | --- |
| `AllowSystemFontEmbedding` | `ResourcePolicy.AllowSystemFontEmbedding` or `CreateTrustedHost()` |
| Markdown `IncludeLocalImages` | `IncludeImages` plus `ResourcePolicy.AllowLocalFileAccess` |
| Markdown `IncludeDataUriImages` | `IncludeImages` plus `ResourcePolicy.AllowDataUris` |

Profiles no longer change trust. Markdown text-only/lightweight profiles only change image participation, and Excel profiles reset their complete profile-owned option set on every application.

Word `IncludePageNumbers` and Excel `IncludeSheetHeadings` now default to `false`; set either to `true` when synthetic visible labels are desired. PowerPoint removed `UseSharedVisualSnapshot`: full-slide PDF always uses its hyperlink-capable native PDF renderer, while PNG/SVG/HTML review and thumbnails use the shared visual snapshot. OneNote now accepts one `OneNotePdfSaveOptions` object and returns explicit semantic-projection diagnostics through `ToPdfDocumentResult()`.

### Migration checklist

- Replace aliases with the canonical names; do not add consumer-side compatibility shims.
- Replace option-owned diagnostics with operation results.
- Use `ToBytes`/`ToStream` for in-memory output and `Save`/`SaveAs{Format}` for destinations.
- Await remote resource resolution and real file/stream I/O; keep pure conversion synchronous.
- Parse HTML into `HtmlConversionDocument` before projecting it to another format.
- Build Reader instances with explicit typed handlers.
- Import shared colors, fonts, images, lifecycle options, and export results from `OfficeIMO.Drawing`.
- Replace image `WithDpi(...)` / `ForHighResolution(...)` with `AtDpi(...)` / `ForPrint(...)`.
- Choose an explicit image file-conflict policy when replacement or unique naming is intended.
- Replace Excel-specific font fallback codes with `OfficeImageExportDiagnosticCodes.FontSubstituted`.
- Use streaming/payload-free batch APIs for production-size page, slide, sheet, chapter, or message exports.
- Treat this as one coordinated package upgrade because old and new surface names are not supported side by side.
