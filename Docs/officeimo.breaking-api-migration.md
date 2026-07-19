# OfficeIMO breaking API migration

This cleanup ships as one coordinated `2.0.0` release across every supported OfficeIMO package. The shared version marks a single compatibility boundary: applications should upgrade their OfficeIMO package set together instead of mixing `1.x` and `2.x` packages.

This release is a coordinated breaking cleanup across the OfficeIMO solution. It removes compatibility aliases, duplicate infrastructure, misleading async methods, and option-owned operation state. Consumers should migrate to the canonical APIs below instead of recreating removed names in wrappers.

## Package architecture

`OfficeIMO.Drawing` remains the small shared foundation for document packages. It already owns the cross-format types required by Word, Excel, PowerPoint, Visio, HTML, PDF, fonts, colors, images, charts, lifecycle options, stream helpers, and export results. There is no additional `OfficeIMO.Core` package and no `.Drawing` to `.Core` rename in this release.

The ownership rules are:

- native format packages own parsing, loading, editing, validation, and serialization for their format;
- adapter packages project one native model into another and do not implement a second parser or document brain;
- `OfficeIMO.Reader.Core` owns normalized read orchestration and contracts, while format-specific Reader packages register typed handlers;
- `OfficeIMO.Html` owns the canonical HTML source model, resource policy, media filtering, and render scene;
- shared colors, fonts, images, stream contracts, lifecycle options, and image export results live in `OfficeIMO.Drawing`;
- `OfficeIMO.Security` owns neutral CMS/X.509/RFC 3161 operations, while each format package owns only its signed-artifact orchestration.

The former compiled `OfficeIMO.Shared` implementation layer is gone. `OfficeIMO.SharedSource` is source-only, and reusable runtime behavior has an explicit owner.

## Persistence lifecycle

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

## Stream ownership

Caller-owned streams are never disposed by OfficeIMO.

- A seekable input is read from the beginning and restored to its original position.
- A non-seekable input is read from its current position to the end.
- A returned stream is new, seekable, and positioned at zero.
- A stream retained as a mutable document destination must be writable and seekable so a later parameterless `Save()` can replace the complete artifact.
- A one-time `Save(stream)` does not silently redirect future parameterless saves unless that document's create/load lifecycle explicitly associates the stream.

These rules are shared across synchronous and asynchronous reads. Cancellation restores a seekable caller stream before the cancellation escapes.

## Async contract

`Async` means the operation performs asynchronous I/O or asynchronous external resource resolution. Pure parsing, model projection, byte generation, and in-memory formatting remain synchronous.

Remote image and stylesheet operations are async-only. For example:

```csharp
HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
HtmlToWordResult converted = await source.ToWordDocumentResultAsync(options, cancellationToken);
```

The synchronous HTML-to-Word API is deliberately offline-only. It accepts embedded and local resources allowed by the operation policy but rejects an import that would perform HTTP I/O.

Removed fake-async methods include in-memory Markdown/HTML/RTF conversions, byte-returning conversion wrappers, `RtfDocument.ReadAsync(string)`, and `RtfDocument.LoadAsync(byte[])`. Use the synchronous conversion, or use `LoadAsync`, `SaveAsync`, and `SaveAs{Format}Async` when the source or destination performs real I/O.

## Conversion results and diagnostics

Reusable option objects contain configuration only. They no longer retain `LastSaveReport`, `LastSaveDiagnostics`, `ConversionReport`, or `Warnings` from a previous operation.

Structured conversion results consistently provide:

- `Value` for the converted model or encoded output;
- `Report` for diagnostics and fidelity evidence;
- `HasLoss` when the conversion simplified or omitted content;
- `RequireValue()` and `RequireNoLoss()` where failing fast is useful.

The canonical PDF result method is `ToPdfDocumentResult()`. Source-explicit methods include `ToPdfDocumentFromMarkdownResult()`, `ToPdfDocumentFromRtfResult()`, and `ToWordDocumentFromPdfResult()`.

`SaveAsPdf` now returns structured save evidence across Word, Excel, PowerPoint, HTML, Markdown, and RTF PDF adapters. `ToPdf()` remains the direct encoded-byte convenience API. Launching or opening a generated PDF is application behavior and is not part of saving.

RTF bridges use `RtfConversionResult<T>`. PDF save attempts expose their report, warnings, warning state, and write outcome rather than mutating the conversion options.

## Image export

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

## HTML source ownership

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

## Reader ownership

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

## Theme ownership

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

## Image export diagnostics

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

## Canonical member names

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
| PDF `ToWordResult()` | `ToWordDocumentFromPdfResult()` |
| `PdfSaveResult.ConversionWarnings` | `Warnings` and `Report` |
| `RtfDocument.ToMemoryStream()` | `ToStream()` |
| `RtfDocument.ToHtmlMemoryStream()` | `ToHtmlStream()` |
| `ToRtfMemoryStream()` | `ToRtfStream()` |
| `SavePdfAsWord()` / `SavePdfAsRtf()` | `SaveAsWordFromPdf()` / `SaveAsRtfFromPdf()` |
| `SavePdfTablesAsExcel/Word/PowerPoint()` | `SaveAsExcel()` / `SaveAsWordDocument()` / `SaveAsPowerPoint()` |
| `WordHelpers.ConvertDotXtoDocX(...)` | `ConvertDotxToDocx(...)` |
| `EmailDocument.WriteToBytes()` | `EmailDocument.ToBytes()` |

Generic file-copy, file-lock probing, duplicate color helpers, public internal save writers, and other APIs with no useful Office document contract were removed rather than renamed.

## PDF converter trust and fidelity defaults

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

## Migration checklist

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
