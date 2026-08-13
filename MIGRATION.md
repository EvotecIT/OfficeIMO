# Upgrading OfficeIMO

This guide contains version-to-version changes that require application code, package references, or configuration to change. It is not a release history or a second API manual.

- Use [GitHub Releases](https://github.com/EvotecIT/OfficeIMO/releases) for release notes and downloadable artifacts.
- Use the root and package READMEs for the current public API.
- Use support matrices for current coverage and limits.
- Use this guide when an upgrade no longer compiles or changes an existing workflow.

OfficeIMO 3.2 is a coordinated package-ownership cleanup. Upgrade every OfficeIMO package in an application to the same `3.2.x` version and perform a clean restore after changing versions.

## OfficeIMO 3.2: bounded object-table and MHTML ingestion defaults

Direct `ObjectFlattener.Flatten`, `GetPaths`, and `ResolvePaths` calls and object-backed Excel and PowerPoint tables now default to at most 16,384 projected columns. Object tables additionally default to at most 1,000,000 cells, including the header row. Set `ObjectFlattenerOptions.MaxColumns` or `MaxCells` explicitly when a trusted workflow needs a different application-level limit. Excel output remains constrained by worksheet dimensions. The object and explicit-binding `PowerPointSlide.AddTable` overloads apply stricter format-safety ceilings of 1,024 columns and 100,000 cells; split larger trusted datasets into multiple tables.

`SheetComposer.TableFrom(DataTable)` previously allowed a fixed-schema table to proceed up to Excel's worksheet dimensions without applying `MaxRows` or `MaxCells`. It now defaults to at most 2,000,000 cells, including the header row, and validates both limits before writing worksheet content. Existing trusted reports above that size must set an explicit bounded override, for example `configure: options => options.MaxCells = requiredCellCount`. The separate Excel worksheet row and column limits cannot be raised.

The aggregate Reader now limits MHT/MHTML input to 64 MiB by default through `OfficeDocumentReaderBuilderMhtmlExtensions.DefaultMaxInputBytes`. Applications can lower or raise that limit by passing `new ReaderOptions { MaxInputBytes = ... }` to the read operation after registering `AddMhtmlHandler()`; use a larger value only for trusted archives with an application-owned resource policy.

## OfficeIMO 3.2: bounded SVG raster fallback

`OfficeVisualSvgPolicy.RasterizeWhenNeeded` now validates the complete rendered SVG expansion before calling the ChartForgeX rasterizer. Valid SVG can therefore throw `InvalidOperationException` when local clip, mask, filter, marker, or pattern references come from a stylesheet that the safety traversal cannot account for precisely. Raising `MaximumSvgElements` does not bypass this conservative rejection.

For trusted generated SVG, replace stylesheet-backed local references with equivalent presentation attributes or inline declarations so each rendered reference can be charged directly. Applications accepting external SVG should handle the exception as an unsupported or over-complex input. Use `PreserveVector` when retaining the partial Office drawing is acceptable, or `RequireVector` when unsupported vector content must fail without raster fallback.

## OfficeIMO 3.2: bounded Visio shape-data projection

Visio document projection now retains at most 200 Shape Data rows per page by default. The limit applies to projected tables, Markdown, and block text, preventing untrusted diagrams from causing unbounded row materialization.

Trusted workflows that need more rows can set an explicit limit:

```csharp
OfficeDocumentModel model = document.ToOfficeDocumentModel(
    options: new VisioDocumentProjectionOptions { MaxTableRows = 5_000 });
```

Use `int.MaxValue` only when trusted input must preserve the former effectively unbounded behavior and the application enforces its own resource policy.

## OfficeIMO 3.2: complete image validation at ingestion boundaries

Excel file and URL image methods now validate the complete bounded payload instead of trusting a filename extension or image header. They throw `ArgumentException` for truncated, corrupt, or unsupported content that older versions could package. When a valid image has a misleading filename extension or remote content type, OfficeIMO uses the format detected from its payload.

Applications that intentionally import opaque package bytes can keep using the byte-array `AddImage(...)` overload with an explicit content type. That is the low-level package path; it does not turn invalid content into a renderable image. Use `OfficeImageReader.TryValidateContent(...)` before ingestion when the application needs to report validation failures itself.

Direct `OfficeImageExportResult` construction now applies the same complete-content check. A recognizable header is no longer enough: truncated, corrupt, undecodable, or dimension-mismatched bytes throw `ArgumentException`. Call `OfficeImageReader.TryValidateContent(...)` before construction when the application needs to handle invalid output without an exception.

Word/RTF result conversions now validate and inventory images across the document body, section headers and footers, fields, revisions, notes, and comments. Images that the target format cannot emit remain in the conversion report instead of disappearing silently. Use `ToRtfDocumentResult(...)` or `ToWordDocumentResult(...)`, inspect `Report`, and call `RequireNoLoss()` when omitted image content must stop the workflow.

Word-to-ODT and PowerPoint-to-ODP conversion now preserves only images that pass `OfficeImageReader.TryValidateContent(...)`. Valid payloads with misleading extensions are stored under the detected format; corrupt, truncated, unsupported, and general WebP payloads outside OfficeIMO's managed decoder subset are omitted and reported as unsupported. Each conversion validates at most 256 image payloads, 128 MiB of aggregate encoded image data, and 100 million aggregate raster pixels; later images are omitted and reported once a ceiling is reached. Split larger trusted documents before conversion when every image must be retained. Inspect the returned `OdfConversionResult<T>.Report`, call `RequireNoLoss()`, or set the conversion option `LossPolicy` when image loss must fail the conversion.

Direct ODT and ODP byte-array `AddImage(...)` methods now validate complete image content. When the filename has a recognized image extension, it must agree with the detected format. These methods throw `ArgumentException` for corrupt, truncated, or mislabeled payloads instead of creating an OpenDocument package entry whose media type does not match its bytes.

## OfficeIMO 3.2: bounded PDF text clipping

PDF reading now accepts at most 4,096 pending glyph clipping paths in one text object. Older versions continued accumulating larger clipping-mode text runs; current versions throw `PdfReadLimitException` with `Kind == PdfReadLimitKind.TextClippingPaths` before that accumulation can exhaust memory. This safety ceiling is not configurable through `PdfReadLimits`. Applications that accept external PDFs should handle the exception as an unsupported or over-complex input. Trusted producers must simplify the clipping text or split it into smaller text objects before OfficeIMO reads the file.

## OfficeIMO 3.2: one PDF authoring and operation model

`PdfDocument` no longer duplicates every heading, paragraph, table, image, form,
and page-layout method on the root object. New PDFs use the composition callback;
an existing generated document can receive another composition through
`Compose(...)`.

```csharp
// OfficeIMO 3.1
PdfDocument.Create(options)
    .H1("Service report")
    .Paragraph(paragraph => paragraph.Text("Ready"))
    .Save("report.pdf");

// OfficeIMO 3.2
PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Service report")
        .Paragraph(paragraph => paragraph.Text("Ready"))), options)
    .Save("report.pdf");
```

Document-wide settings still belong in `PdfOptions`. Page-scoped headers,
footers, backgrounds, watermarks, and layout use `pdf.Page(page => ...)`.
Reusable content stays on `PdfItemCompose`, so adapters and applications share
one authoring vocabulary.

Specialized existing-document operations now use capability objects:

| OfficeIMO 3.1 usage | OfficeIMO 3.2 replacement |
| --- | --- |
| `document.Encrypt(options)` | `document.Security.Encrypt(options)` |
| `document.ValidateSignatures(provider)` | `document.Security.ValidateSignatures(provider)` |
| `document.PlanRedactions(areas)` | `document.Redactions.Plan(areas)` |
| `document.ApplyRedactions(plan)` | `document.Redactions.Apply(plan)` |
| `document.AnalyzeOptimization()` | `document.Optimization.Analyze()` |
| `document.Optimize(profile)` | `document.Optimization.Apply(profile)` |
| `document.CompareVisual(actual)` | `document.Proof.CompareVisual(actual)` |
| `document.AssessRewritePreservation(rewritten)` | `document.Proof.AssessRewritePreservation(rewritten)` |

`Pages`, `Read`, `Forms`, `Attachments`, `Bookmarks`, `Annotations`, and `Stamp`
keep their existing capability-object shape. The former static implementation
engines remain internal; applications should not replace the removed root
methods with calls to those engines.

## OfficeIMO 3.2: bounded RTF reads by default

`RtfReadOptions` now defaults to the bounded OfficeIMO profile. Embedded objects and file-table references are not materialized, hyperlink fields are restricted to web and mail schemes, and byte, character, token, group, payload, image, object, and semantic-block limits apply.

Applications that intentionally rely on the former permissive behavior for trusted files must opt in:

```csharp
RtfReadResult result = RtfDocument.Load(
    "trusted-legacy.rtf",
    RtfReadOptions.CreateCompatibilityProfile());
```

Do not use the compatibility profile for uploads or other untrusted inputs. Lossless byte output from character-only reads now fails when the source cannot be represented exactly; use byte, stream, or file input when exact original bytes are required.

## OfficeIMO 3.2: bounded LaTeX byte input by default

LaTeX file and stream loading now rejects encoded input larger than 64 MiB before decoding, independently of the existing decoded-character limit. Applications that intentionally load larger trusted documents must raise or disable the byte limit explicitly:

```csharp
LatexParseResult result = LatexDocument.Load(
    "trusted-large-document.tex",
    new LatexParseOptions { MaximumInputBytes = null });
```

Keep the default for uploads and other untrusted input. Set `MaximumInputBytes` to a larger finite value when the application has a known document-size ceiling; use `null` only for a trusted source with a separate resource policy.

### PDF OCR provider coordinates

`PdfOcrRequest.PageWidth` and `PageHeight` now describe the rendered visual page
after applying the crop box and page rotation. In earlier versions they exposed
the unrotated logical dimensions even though `Png`, `PixelWidth`, and
`PixelHeight` represented the rendered page. For pages rotated 90 or 270
degrees, the point dimensions are therefore swapped.

Existing `IPdfOcrProvider` implementations should map their pixel-space
`PdfOcrWord` results against these visual dimensions. Providers that cached or
recomputed unrotated media-box dimensions should instead use the request's
`PageWidth`, `PageHeight`, and `Scale`, which now describe the same visual
coordinate space as the supplied PNG.

## OfficeIMO 3.2: neutral conversion model

Direct format conversion no longer uses Reader as its intermediate ownership
layer. `OfficeIMO.Core` now contains the dependency-free `OfficeDocumentModel`,
source formats project into that model, and destination packages own their
output policy.

| OfficeIMO 3.1 usage | OfficeIMO 3.2 replacement |
| --- | --- |
| `OfficeIMO.Reader.ReaderPdfProjectionOptions` | `OfficeIMO.Pdf.PdfProjectionOptions` |
| `ReaderPdfPagePolicy` | `PdfProjectionPagePolicy` |
| `ReaderPdfAssetPolicy` | `PdfProjectionAssetPolicy` |
| `ReaderPdfLinkPolicy` | `PdfProjectionLinkPolicy` |
| `ReaderPdfFormPolicy` | `PdfProjectionFormPolicy` |
| Reader options passed through `VisioPdfSaveOptions` | `VisioDocumentProjectionOptions` for source projection and `PdfProjectionOptions` for PDF output |

`OfficeIMO.Visio.Pdf` now depends only on `OfficeIMO.Core`, `OfficeIMO.Visio`,
and `OfficeIMO.Pdf`; it no longer installs `OfficeIMO.Reader.Visio` or
`OfficeIMO.Reader.Pdf`. Existing Reader-result-to-PDF calls remain available
through a thin `OfficeIMO.Reader.Pdf` compatibility bridge, but the projection
implementation and options belong to `OfficeIMO.Pdf`.

## OfficeIMO 3.2: explicit HTML and MHTML bridges

`OfficeIMO.Html` is now the lean HTML engine. It no longer makes ordinary HTML, Markdown, Office-conversion, Reader, or EPUB applications acquire Email or RTF packages. Cross-format behavior moved to packages whose names describe both sides of the operation.

| Existing workflow | 3.2 package and code change |
| --- | --- |
| HTML to/from RTF | Add `OfficeIMO.Html.Rtf`. Existing APIs remain in the `OfficeIMO.Html` namespace. |
| Load or save MHT/MHTML | Add `OfficeIMO.Mhtml` and replace `using OfficeIMO.Html;` for `MhtmlDocument` or `MhtmlResource` with `using OfficeIMO.Mhtml;`. |
| Render an `EmailDocument` to images | Add `OfficeIMO.Email.Image`. Existing APIs remain in the `OfficeIMO.Email` namespace. |
| Convert MHT/MHTML to PDF | Add `OfficeIMO.Mhtml.Pdf` plus `OfficeIMO.Mhtml`; use the `OfficeIMO.Mhtml` extension namespace. Plain HTML/PDF remains in `OfficeIMO.Html.Pdf`. |
| Register HTML with Reader | `AddHtmlHandler()` now registers `.html`, `.htm`, and `.xhtml` only. |
| Register MHT/MHTML with Reader | Reference `OfficeIMO.Reader.Email`, import `OfficeIMO.Reader.Email`, and call `AddMhtmlHandler()`. `AddEmailHandlers()` and `OfficeIMO.Reader.All` include it automatically. |
| Read EPUB through Reader | No source change. `OfficeIMO.Reader.Epub` still reuses the HTML projection but no longer receives Email, RTF, or MHTML transitively. |

No `OfficeIMO.Html.Core`, separate document-model package, or `OfficeIMO.Reader.Mhtml` package was introduced. The base HTML and Reader APIs stay focused; optional bridges carry the extra dependency edges.

## OfficeIMO 3.1

OfficeIMO 3.1 was a coordinated breaking release. The sections below describe upgrades from 3.0 to aligned `3.1.x` packages.

## Start here: most OfficeIMO.Word applications

If an application references `OfficeIMO.Word` and uses the high-level Word object
model, start here. The usual create, load, edit, and save workflow keeps the same
shape. Most migration work is replacing enum or type names reported by the
compiler; applications do not need to rewrite ordinary paragraph, table, image,
header, footer, or section workflows.

Use this upgrade sequence:

1. Upgrade `OfficeIMO.Word` and every other OfficeIMO package in the application
   to the same `3.1.x` version.
2. If the project explicitly references the `OfficeIMO.Drawing` package, replace
   that package reference with `OfficeIMO.Core`. A project that references only
   `OfficeIMO.Word` receives Core through the Word package dependency.
3. Delete stale `bin` and `obj` output, restore packages, and build the application.
4. Fix compiler errors using the Word and shared replacement tables below. Remove
   `DocumentFormat.OpenXml.Wordprocessing` imports that were present only for
   high-level OfficeIMO enum values.
5. Run the application's document tests. Review the focused behavior notes only
   for APIs the application actually uses, especially table layout, low-level
   style/page settings, signatures, or converters.

A typical enum migration is local:

```csharp
// OfficeIMO 3.0
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

title.SetAlignment(JustificationValues.Center);
document.AddSection(SectionMarkValues.Continuous);
```

```csharp
// OfficeIMO 3.1
using OfficeIMO.Word;

title.SetAlignment(WordParagraphAlignment.Center);
document.AddSection(WordSectionBreakType.Continuous);
```

Estimate the likely migration from the APIs the application uses:

| Application shape | Likely work |
| --- | --- |
| `OfficeIMO.Word` with the regular document object model | Mostly compiler-guided enum and type replacements. |
| Word fluent APIs without storing concrete builder types | The `AsFluent()` entry point remains; rename any builder types referenced explicitly. |
| Word code using Open XML values as high-level options | Replace them with the corresponding `Word*` or shared `Office*` enums. |
| Word code receiving Open XML styles, page sizes, note settings, or page-number elements from OfficeIMO | Move to the OfficeIMO-owned definitions and settings described under [cross-format contract and naming cleanup](#cross-format-contract-and-naming-cleanup). |
| Word plus PDF, HTML, Markdown, RTF, or Google Docs adapters | Also replace format-specific conversion diagnostics and policies with the shared `OfficeConversion*` contracts. |
| Word, Excel, and PowerPoint in one application | Expect more renames, but fewer ambiguous types afterward because genuinely shared contracts now have one `Office*` name. |
| Removed Excel reader APIs or the old PowerPoint fluent/composer surface | Follow the dedicated [CSV and Excel tabular reads](#csv-and-excel-tabular-reads) or [PowerPoint lifecycle, composition, and inspection](#powerpoint-lifecycle-composition-and-inspection) sections; these are the migrations most likely to require workflow changes. |

The rest of this guide is a lookup reference. An application does not need to
apply sections for packages or features it does not use.

## Shared foundation package: Drawing to Core

The zero-dependency shared package, project, and assembly have been renamed from
`OfficeIMO.Drawing` to `OfficeIMO.Core`:

```xml
<PackageReference Include="OfficeIMO.Core" Version="3.1.0" />
```

This is a rename of the existing shared foundation, not a split into another
dependency layer. Drawing originally absorbed the shared primitives so Word,
Excel, PowerPoint, Visio, PDF, and the conversion packages would not require a
Core package that then depended on a separate Drawing package. Over time that
assembly also became the owner of lifecycle, package-security, embedded-payload,
and neutral data-mapping contracts, so `OfficeIMO.Drawing` no longer described
the package honestly.

Actual drawing types such as `OfficeColor`, `OfficeShape`, and
`OfficeRenderingProfile` remain in the `OfficeIMO.Drawing` namespace. Security
provider APIs remain in `OfficeIMO.Security`. Cross-document lifecycle and
package contracts, compatibility models, capability catalogs, and conversion
reports move to the root `OfficeIMO` namespace. Neutral object-flattening and
row-mapping contracts use `OfficeIMO.Data`:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| Package/assembly `OfficeIMO.Drawing` | Package/assembly `OfficeIMO.Core` |
| `OfficeIMO.Drawing.DocumentAccessMode` | `OfficeIMO.DocumentAccessMode` |
| `OfficeIMO.Drawing.DocumentPersistenceMode` | `OfficeIMO.DocumentPersistenceMode` |
| `OfficeIMO.Drawing.DocumentCreateOptions` | `OfficeIMO.DocumentCreateOptions` |
| `OfficeIMO.Drawing.DocumentLoadOptions` | `OfficeIMO.DocumentLoadOptions` |
| `OfficeIMO.Drawing.OfficePackageSecurityOptions` and related package contracts | Root `OfficeIMO` namespace |
| `OfficeIMO.Drawing.OfficeFormatDescriptor`, compatibility models, and capability catalogs | Root `OfficeIMO` namespace |
| `OfficeIMO.Drawing.IOfficeConversionReport` | `OfficeIMO.IOfficeConversionReport` |
| `OfficeIMO.Drawing.ObjectFlattener`, `ObjectFlattenerOptions`, and `CollectionColumnMapping` | `OfficeIMO.Data` namespace |
| `OfficeIMO.Drawing.HeaderCase`, `NullPolicy`, and `CollectionMode` | `OfficeIMO.Data` namespace |

Replace the package reference and add `using OfficeIMO;` where lifecycle or
package, compatibility, capability, or conversion-report contracts are used.
Add `using OfficeIMO.Data;` for row mapping and object flattening. Keep
`using OfficeIMO.Drawing;` for actual drawing, color, image, chart, font, and
rendering APIs.

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
| Cryptographic operations reached through a format package's transitive `OfficeIMO.Security` dependency | Install `OfficeIMO.Security` explicitly and pass `OfficeSecurityProvider.Default` through the format API. |

There is no replacement umbrella OpenDocument PDF package. The focused packages keep Word, Excel, and PowerPoint dependencies out of applications that do not use those routes.

### OpenXML value types are now OfficeIMO enums

OpenXML SDK 3 changed schema value types such as `SectionMarkValues` from CLR enums into generated value structs. Those structs work as OpenXML serialization values, but they are poor high-level API contracts: PowerShell cannot bind them like enums, and applications become coupled to an SDK implementation detail.

OfficeIMO 3.1 public APIs therefore accept and return OfficeIMO-owned CLR enums. This is an intentional breaking change. Remove `DocumentFormat.OpenXml` imports that were used only to select a high-level OfficeIMO option, and use the corresponding `Word*`, `PowerPoint*`, or `Excel*` enum. Direct package-element editing can still use OpenXML SDK types at that low level.

Common Word changes are:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `document.AddSection(SectionMarkValues.Continuous)` | `document.AddSection(WordSectionBreakType.Continuous)` |
| `paragraph.AddBreak(BreakValues.Page)` | `paragraph.AddBreak(WordBreakType.Page)` |
| `paragraph.SetUnderline(UnderlineValues.Double)` | `paragraph.SetUnderline(WordUnderlineStyle.Double)` |
| `paragraph.SetAlignment(JustificationValues.Center)` | `paragraph.SetAlignment(WordParagraphAlignment.Center)` |
| `section.GetOrCreateHeader(HeaderFooterValues.First)` | `section.GetOrCreateHeader(WordHeaderFooterType.First)` |
| `PageOrientationValues`, `WordPageOrientation`, `ExcelPageOrientation`, or `PdfPageOrientation` | `OfficePageOrientation` |
| `NumberFormatValues` | `WordNumberFormat` |
| `BorderValues` | `WordBorderStyle` |
| `HighlightColorValues` | `WordHighlightColor` |
| `ShadingPatternValues` | `WordShadingPattern` |
| `TabStopValues` / `TabStopLeaderCharValues` | `WordTabAlignment` / `WordTabLeader` |
| `TableLayoutValues` / `TableOverlapValues` | `WordTableLayoutMode` / `WordTableOverlap` |
| `TableRowAlignmentValues` / `TableWidthUnitValues` | `WordTableAlignment` / `WordTableWidthUnit` |
| `TableVerticalAlignmentValues` / `TextDirectionValues` | `WordTableVerticalAlignment` / `WordTextDirection` |
| `HorizontalRelativePositionValues` / `VerticalRelativePositionValues` | `WordHorizontalRelativePosition` / `WordVerticalRelativePosition` |
| `HorizontalAnchorValues` / `VerticalAnchorValues` | `WordTableHorizontalAnchor` / `WordTableVerticalAnchor` |
| `HorizontalAlignmentValues` / `VerticalAlignmentValues` | `WordTableHorizontalAlignment` / `WordTableVerticalPositionAlignment` |
| `VerticalPositionValues` / `VerticalTextAlignmentValues` | `WordVerticalTextPosition` / `WordVerticalCharacterAlignment` |
| `DocumentProtectionValues` | `WordDocumentProtectionType` |
| `FootnotePositionValues` / `EndnotePositionValues` / `RestartNumberValues` | `WordFootnotePosition` / `WordEndnotePosition` / `WordNoteNumberRestart` |
| `LevelJustificationValues` / `LevelSuffixValues` | `WordListLevelAlignment` / `WordListLevelSuffix` |
| `ShapeTypeValues` / `BlipCompressionValues` / `BlackWhiteModeValues` | `OfficePresetShapeType` / `WordImageCompressionQuality` / `WordImageBlackWhiteMode` |
| `WordImage.BlackWiteMode` | `WordImage.BlackWhiteMode` (the misspelled member remains as an obsolete forwarding alias) |
| chart `BarDirectionValues`, `BarGroupingValues`, and `LegendPositionValues` | `WordChartBarDirection`, `WordChartBarGrouping`, and `OfficeChartLegendPosition` |

PowerPoint uses the same rule:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `AddSlideWithLayoutType(...)` / `SetLayoutWithType(...)` / `GetLayoutIndexWithType(...)` | `AddSlide(...)` / `SetLayout(...)` / `GetLayoutIndex(...)` with `PowerPointSlideLayoutType` |
| `SlideLayoutValues` | `PowerPointSlideLayoutType` |
| `ShapeTypeValues` | `OfficePresetShapeType` |
| `TextAlignmentTypeValues` / `TextAnchoringTypeValues` / `TextVerticalValues` | `PowerPointTextAlignment` / `PowerPointTextVerticalAlignment` / `PowerPointTextDirection` |
| `TextAutoNumberSchemeValues` / `TextUnderlineValues` | `PowerPointNumberingScheme` / `PowerPointUnderlineStyle` |
| `LineEndValues` / `LineEndLengthValues` / `LineEndWidthValues` | `OfficeLineMarkerKind` / `PowerPointLineEndLength` / `PowerPointLineEndWidth` |
| `PresetLineDashValues` / `RectangleAlignmentValues` | `PowerPointLineDashStyle` / `PowerPointRectangleAlignment` |
| `PlaceholderValues` / `PlaceholderSizeValues` / `DirectionValues` | `PowerPointPlaceholderType` / `PowerPointPlaceholderSize` / `PowerPointPlaceholderDirection` |
| `SlideSizeValues` | `PowerPointSlideSizeType` |
| chart `BuiltInUnitValues`, `CrossBetweenValues`, `CrossesValues`, and `TickLabelPositionValues` | `OfficeChartDisplayUnit`, `OfficeChartAxisCrossBetween`, `OfficeChartAxisCrossingPosition`, and `OfficeChartAxisTickLabelPosition` |
| chart `DataLabelPositionValues`, `GroupingValues`, `LegendPositionValues`, `MarkerStyleValues`, and `TrendlineValues` | the corresponding `PowerPointChart*` enum |

Excel replacements include:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `BorderStyleValues` / `HorizontalAlignmentValues` / `VerticalAlignmentValues` | `ExcelBorderStyle` / `ExcelHorizontalAlignment` / `ExcelVerticalAlignment` |
| `CellValues` / `UnderlineValues` / `VerticalAlignmentRunValues` | `ExcelCellValueType` / `ExcelUnderlineStyle` / `ExcelVerticalTextAlignment` |
| `ConditionalFormatValues` / `ConditionalFormattingOperatorValues` / `TimePeriodValues` | `ExcelConditionalFormatType` / `ExcelConditionalFormattingOperator` / `ExcelConditionalTimePeriod` |
| `DataValidationErrorStyleValues` / `DataValidationOperatorValues` | `ExcelDataValidationErrorStyle` / `ExcelDataValidationOperator` |
| `IconSetValues` | `ExcelIconSet` |
| `SparklineTypeValues` | `ExcelSparklineType` |
| `TotalsRowFunctionValues` | `ExcelTableTotalsFunction` |
| `DataConsolidateFunctionValues` / `FieldSortValues` / `GroupByValues` | `ExcelPivotDataFunction` / `ExcelPivotFieldSort` / `ExcelPivotGroupBy` |
| `PivotFilterValues` / `PivotTableAxisValues` / `ShowDataAsValues` | `ExcelPivotFilterType` / `ExcelPivotTableAxis` / `ExcelPivotShowDataAs` |
| chart `BuiltInUnitValues`, `CrossBetweenValues`, `CrossesValues`, and `TickLabelPositionValues` | `OfficeChartDisplayUnit`, `OfficeChartAxisCrossBetween`, `OfficeChartAxisCrossingPosition`, and `OfficeChartAxisTickLabelPosition` |
| chart `DataLabelPositionValues`, `LegendPositionValues`, `MarkerStyleValues`, and `TrendlineValues` | the corresponding `ExcelChart*` enum |

The OpenXML member named `ShowDataAsValues.PercentOfRaw` serializes the token `percentOfRow`. OfficeIMO exposes the corrected spelling `ExcelPivotShowDataAs.PercentOfRow`.

### Cross-format contract and naming cleanup

OfficeIMO applications commonly import Word, Excel, PowerPoint, PDF, and one or
more converters together. Format-owned types therefore carry their format name,
while contracts with the same meaning in several formats live once in
`OfficeIMO.Core` with an `Office*` name. This avoids ambiguous `using`
directives and prevents converters from defining parallel policy enums.

Shared replacements include:

| OfficeIMO 3.0 or early 3.1 preview | OfficeIMO 3.1 |
| --- | --- |
| Format-specific conversion loss and destination-conflict enums | `OfficeConversionLossPolicy` / `OfficeConversionFileConflictPolicy` |
| Format-specific conversion diagnostic category, severity, and failure enums | `OfficeConversionDiagnosticCategory`, `OfficeConversionDiagnosticSeverity`, and `OfficeConversionFailureReason` |
| Format-specific feature-support enums | `OfficeFeatureSupportLevel` |
| Open XML SDK `OpenSettings` in Word, Excel, or PowerPoint load options | `OfficeOpenXmlLoadSettings` |
| Open XML SDK `FileFormatVersions` in public validation APIs | `OfficeOpenXmlFileFormatVersion` |
| Open XML SDK validation errors | `OfficeOpenXmlValidationError` |
| Word/Excel/PowerPoint copies of common chart positions, markers, line ends, and preset shapes | `OfficeChart*`, `OfficeLineMarkerKind`, and `OfficePresetShapeType` from `OfficeIMO.Core` |
| Word/PowerPoint image-part format enums | `OfficeImageFormat` from `OfficeIMO.Core` |

Format-specific public type renames include:

| OfficeIMO 3.0 or early 3.1 preview | OfficeIMO 3.1 |
| --- | --- |
| Word `ApplicationProperties` / `BuiltinDocumentProperties` | `WordApplicationProperties` / `WordBuiltinDocumentProperties` |
| Excel `ApplicationProperties` / `BuiltinDocumentProperties` | `ExcelApplicationProperties` / `ExcelBuiltinDocumentProperties` |
| `CapsStyle` / `CompatibilityMode` / `CoverPageTemplate` | `WordCapsStyle` / `WordCompatibilityMode` / `WordCoverPageTemplate` |
| `CustomImagePartType` / `WordImagePartType` / `ImageFillMode` / `WrapTextImage` | `OfficeImageFormat` / `WordImageFillMode` / `WordImageTextWrapping` |
| `ShapeType` / `SmartArtType` | `WordShapeType` / `WordSmartArtType` |
| `TableOfContentStyle` / `TargetFrame` / `TextMatchType` | `WordTableOfContentsStyle` / `WordHyperlinkTargetFrame` / `WordTextMatchType` |
| Word `DocumentCleanupOptions` / `PropertyTypes` | `WordDocumentCleanupOptions` / `WordCustomPropertyType` |
| Excel `ExecutionMode` / `ExecutionPolicy` | `ExcelExecutionMode` / `ExcelExecutionPolicy` |
| Excel `HeaderFooterPosition` / `NameValidationMode` | `ExcelHeaderFooterPosition` / `ExcelDefinedNameValidationMode` |
| Excel `TableStyle` | `ExcelTableStyle` |
| Excel `SheetNameValidationMode` / `TableNameValidationMode` / `WorksheetValidationMode` | `ExcelSheetNameValidationMode` / `ExcelTableNameValidationMode` / `ExcelWorksheetValidationMode` |
| PowerPoint `ImagePartType` / `PowerPointImagePartType` | `OfficeImageFormat` |
| `SlideTransition` / `SlideTransitionSpeed` / `TableCellBorders` | `PowerPointSlideTransition` / `PowerPointSlideTransitionSpeed` / `PowerPointTableCellBorders` |
| Excel or PowerPoint format-specific signature mutation policy | `OfficeSignatureMutationPolicy` |
| `OfficeImageExportLossKind`, `HtmlConversionLossKind`, or `WordMarkdownConversionLossKind` | `OfficeConversionLossKind` |
| `GoogleDocsImportMode`, `GoogleSheetsImportMode`, or `GoogleSlidesImportMode` | `GoogleWorkspaceImportMode` |
| `GoogleDocsDiffKind`, `GoogleSheetsDiffKind`, or `GoogleSlidesDiffKind` | `GoogleWorkspaceDiffKind` |

Word high-level APIs no longer expose package elements for ordinary formatting:

| OfficeIMO 3.0 or early 3.1 preview | OfficeIMO 3.1 |
| --- | --- |
| Open XML `Style` arguments and results on paragraph-style APIs | `WordParagraphStyleDefinition` |
| Open XML page-size values | `WordPageSizeDefinition` |
| Open XML `FootnoteProperties`, `EndnoteProperties`, or `PageNumberType` | `WordFootnoteSettings`, `WordEndnoteSettings`, and `WordPageNumberSettings` |
| Open XML numeric wrappers on margins and borders | `uint` / `ushort` values |
| `WordHorizontalAlignmentValues` for text boxes | `WordTextBoxHorizontalAlignment` |
| Generic fluent `HorizontalAlignment` / `VerticalAlignment` | `WordParagraphAlignment` or `WordTableAlignment`, according to the owning builder |
| `WordTableLayoutType`, raw `WordTableLayoutMode`, `WordTable.LayoutType`, or `SetTableLayout(...)` | `WordTableLayoutMode.AutoFit` / `Fixed` and `WordTable.LayoutMode`; use `AutoFitToContents()`, `AutoFitToWindow()`, or `SetFixedWidth(percent)` when width changes are also intended |
| Word fluent `ParagraphBuilder` / Markdown `ParagraphBuilder` | `WordParagraphBuilder` / `MarkdownParagraphBuilder` |
| PDF `TextRun` / Markdown `TextRun` | `PdfTextRun` / `MarkdownTextRun` |

The early 3.1-preview values `AutoFitToContents`, `AutoFitToWindow`, and
`FixedWidth` mixed Word's two layout algorithms with width presets and could not
round-trip unambiguously. `WordTable.LayoutMode` now changes only the real
layout algorithm and preserves table and cell preferred widths. The explicit
width helpers retain their action-oriented behavior. `SetFixedWidth(percent)`
sets the table width only; it no longer rewrites each cell to an equal width.

Excel package-metadata creation methods now return `ExcelPackagePartInfo`
instead of an Open XML SDK package part. Use its `RelationshipId`, `Uri`,
`ContentType`, and `RelationshipType` properties. Applications that intentionally
perform low-level package editing can still start from
`ExcelDocument.OpenXmlDocument`.

`ExcelSheet.BeginNoLock()` is no longer public. Use `ExcelSheet.Batch(...)` to
group worksheet mutations under one write lock.

`OfficeIMO.Markup` now emits the same OfficeIMO-owned enums in generated C#. Its PowerShell starter output uses canonical PSWriteOffice commands and string enum member names so it remains compatible with the packaged module's isolated dependency context. Regenerate previously emitted starter code, review any emitted `TODO` comments for chart-data binding, or replace remaining `DocumentFormat.OpenXml.*Values` arguments with the matching enum above before compiling or running against 3.1.

### Optional security provider

Word, PDF, and Email no longer pull cryptographic packages into applications that only create, read, inspect, or
convert documents. Applications that create or cryptographically validate signatures, or decrypt S/MIME content,
must add the optional provider explicitly:

```powershell
dotnet add package OfficeIMO.Security
```

```csharp
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
```

Pass that provider to the format-owned operation. The format package continues to own byte ranges, package parts,
relationships, signed-content selection, preservation policy, and document mutation. `OfficeIMO.Security` owns CMS,
XML DSig, X.509, and RFC 3161 cryptography.

| OfficeIMO 3.0 call shape | OfficeIMO 3.1 call shape |
| --- | --- |
| `WordDocument.SignPackage(path, certificateOrThumbprint, ...)` | `WordDocument.SignPackage(path, security, certificateOrThumbprint, ...)` |
| `document.ValidateSignatures(options)` | `document.ValidateSignatures(security, options)` |
| `WordDocument.SignMacroProject(path, certificateThumbprint, options)` | `WordDocument.SignMacroProject(path, security, certificateThumbprint, options)` |
| `WordDocument.ValidateMacroProjectSignature(path, options)` | `WordDocument.ValidateMacroProjectSignature(path, security, options)` |
| `new PdfCmsExternalSigner(certificate, ...)` | `new PdfCmsExternalSigner(security, certificate, ...)` |
| `new PdfCmsSignatureCryptographyProvider(options)` | `new PdfCmsSignatureCryptographyProvider(security, options)` |
| `EmailSmime.Verify(document, options)` | `EmailSmime.Verify(document, security, options)` |
| `EmailSmime.Decrypt(document, certificate, options)` | `EmailSmime.Decrypt(document, certificate, security, options)` |

Structural signature inspection and fail-safe mutation policies remain available without the optional package. Excel,
PowerPoint, Visio, OpenDocument, and EPUB currently inspect or safely handle their native signature carriers without
claiming cryptographic validation; they do not require `OfficeIMO.Security` until a provider-backed format adapter is
implemented.

| Route | Focused package | Reverse entry point |
| --- | --- | --- |
| ODT to or from PDF | `OfficeIMO.OpenDocument.Odt.Pdf` | `pdf.ToOdtDocument()` |
| ODS to or from PDF | `OfficeIMO.OpenDocument.Ods.Pdf` | `pdf.ToOdsDocument()` |
| ODP to or from PDF | `OfficeIMO.OpenDocument.Odp.Pdf` | `pdf.ToOdpPresentation()` |

OpenDocument forward PDF results expose the typed OpenDocument projection report in `SourceConversionReports` and PDF-layout warnings in `Report`; `ConversionReports` presents both stages in order. Read OpenDocument mappings from `OdfConversionReport.Mappings` instead of looking for the removed synthetic `ODF_*` PDF warnings. Use `HasLoss` or `RequireNoLoss()` for the end-to-end fidelity gate. AsciiDoc, LaTeX, and semantic OneNote PDF routes use the same ordered source-stage and PDF-stage model.

`HasWarnings` remains the PDF-stage flag because source reports have format-specific diagnostic models. Use `new PdfConversionProofOptions().RequireNoLoss()` when conversion proof must enforce the same end-to-end fidelity rule. `PdfDocumentConversionResult.Warnings` describes the PDF stage only.

### Google Workspace preview options

Google Workspace mutations now require `ExpectedAccount`, an `OperationPolicyProvider`, and an `OperationReceiptSink` on `GoogleWorkspaceSessionOptions`. Credential sources must attach provider-verified identity and grant evidence with `GoogleWorkspaceAccessToken.FromVerifiedCredential`; caller-entered account labels and requested scopes remain unverified and cannot authorize mutations. The built-in service-account source creates this evidence from its signed assertion and token exchange. Google APIs and installed-application callers supply a `GoogleWorkspaceCredentialBindingResolver` backed by provider token evidence. A provider-verified grant set may contain more scopes than one adapter operation requests; the session accepts that subset while the operation policy and receipt remain bound to the adapter's exact required scopes. Raw `StaticAccessTokenCredentialSource` values remain suitable for reads; mutation-capable raw-token applications must use a delegate source after independently verifying the token. Construct low-level mutation transport with `new GoogleWorkspaceHttpTransport(session)` so it can prove that the supplied token and required scope set were acquired by that session; the options-only constructor remains suitable for reads but rejects mutations. Build each policy from the operation context's `RequiredScopes`, `MaxRetryCount`, `MaxRetryElapsedTime`, and `RateLimitPolicy`; the transport snapshots and verifies those values before sending, so a policy cannot claim different scopes or mutable retry behavior. Policy targets preserve allowlisted operation-defining query values, including Drive parent changes, while sensitive and non-semantic query values are redacted or excluded. Adapter calls declare `GoogleWorkspaceMutationKind` independently of their HTTP verb and expose `RevisionPreconditionKind`: return `AdapterExpectedRevision` for payload-enforced Docs or Slides write control and for `ResumableSessionState`, a strong HTTP entity tag for enforced `If-Match`, `ResourceAbsentForCreateRevision` for an adapter-declared create, or `ExplicitlyUnversionedRevision(reason)` together with an accepted, named loss decision when the API has no usable conditional precondition. Resumable Drive session initiation and chunk receipts are actions; the create receipt is emitted only after Google confirms the completed file. Mutation receipts record the semantic mutation kind, selected mechanism, and revision or session state actually enforced. Any guarded mutation that fails before final response headers now records an ambiguous receipt and throws `GoogleWorkspaceAmbiguousMutationException`; reconcile the receipt target and request identifier before retrying. Sync plan items require a target resource and expected revision, and `GoogleWorkspaceSyncPlan.Create` requires the plan policy. Read sync executor decisions from `GoogleWorkspaceSyncItemResult.DecisionReceipt`; actual network mutation receipts continue to arrive through the session receipt sink.

Email store content-search checkpoints are now source- and query-bound serialized envelopes. Replace direct `new EmailStoreContentSearchCheckpoint(offset)` construction with the checkpoint returned by `EmailStoreContentSearchReport.NextCheckpoint`; persisted legacy offset-only values cannot safely resume against a reopened or changed store.

Email store table-page continuation tokens also use a version-2, source-bound envelope. Persisted version-1 `EmailStoreContinuationToken` values are intentionally rejected because they cannot prove that the reopened source is unchanged. Discard those tokens after upgrading and restart the affected table query from its first page.

The completed Google Workspace adapters replace preview booleans that described behavior without proving that the fallback executed. Configure the operation through `UnsupportedFeatures`, `GoogleWorkspaceFidelityPolicy`, the format support catalog, and an executed fallback mode:

| Removed preview option | Current action |
| --- | --- |
| Google Docs `FlattenFloatingContent` | Set `UnsupportedFeatures.FloatingContent` to the required `UnsupportedFeatureMode`, such as `Flatten`, `Rasterize`, `WarnAndSkip`, or `Error`. |
| Google Docs `RasterizeWordCharts` | Set `UnsupportedFeatures.Charts = UnsupportedFeatureMode.Rasterize`; configure bounded rendered-page output through `RasterFallbackImageOptions`. |
| Google Docs `PreserveCommentsViaDriveApi` | Set `Comments = GoogleDocsCommentMode.UnanchoredDriveComments`; use `UnsupportedFeatures.Comments` for content that cannot use that executed route. |
| Google Docs `IncludeHeadersAndFooters`, `IncludeFootnotes`, and `IncludeBookmarksAsNamedRanges` | Supported content is translated according to the code-owned catalog; gate unsupported content with `FidelityPolicy` and `UnsupportedFeatures` instead of disabling a promise-only switch. |
| Google Sheets `IncludeCharts` | Set `UnsupportedFeatures.Charts` to the required executed or fail-fast mode. |
| Google Sheets `IncludePivotTables` | Set `UnsupportedFeatures.PivotTables` to the required executed or fail-fast mode. |
| Google Sheets `IncludeHeaderFooterMetadata` and `TreatPrintLayoutAsDiagnosticOnly` | Set `UnsupportedFeatures.PrintLayout`, normally to `WarnAndSkip` or `Error`. |
| Google Sheets `PreserveUnsupportedFormulasAsText` | Set `Formulas.UnsupportedFormulaMode` to `PreserveWithWarning`, `UseCachedValue`, or `Error`. |

Read the target before replacement. Docs and Slides require the observed API revision; Sheets requires the observed Drive version. The [generated support matrix](https://officeimo.com/docs/google-workspace/support/) is the current capability owner.

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
| Concrete `ExcelDocumentReader` / `ExcelSheetReader` use | `ExcelWorkbookDataReader` returned by `ExcelDocument.OpenDataReader(...)` or `workbook.CreateDataReader(...)` |
| `ExcelSheet.Rows(...)` | `workbook.CreateDataReader(...)` to read the current open workbook (including unsaved edits), `ExcelDocument.OpenDataReader(...)` for an unopened file/stream, or deferred `ExcelSheet.RowsAs<T>(...)` projection |
| `ExcelSheet.RowsObjects(...)`, `RowEdit`, or `CellEdit` | Direct `ExcelSheet` cell APIs such as `CellValue(...)`, `CellFormula(...)`, and `FormatCell(...)` |
| `ExcelSheet.GetUsedRangeA1()` | `ExcelSheet.UsedRangeA1` |

`CsvLoadOptions.Mode`, `CsvLoadMode`, and `CsvDocument.Mode` are no longer
public. Use `CsvDocument.OpenDataReader` for a forward-only read and
`CsvDocument.Load` for a materialized model.

`CsvDocument.Materialize()` is also no longer public because `Load` and `Parse`
always return a materialized document. There is no public streaming-document
state to convert; use `OpenDataReader` when rows should remain forward-only.

### CSV and Excel typed-row cleanup

OfficeIMO 3.1 has one package-native typed-row writer for Excel and one typed-row
projection name for CSV. Replace the removed overlapping entry points as follows:

| Removed 3.1 preview API | Current 3.1 API |
| --- | --- |
| `ExcelDocument.WriteObjects(..., IReadOnlyList<(string Header, Func<T, object?> Selector)>, ...)` | `ExcelDocument.WriteRows(..., headers, (writer, row) => writer.Write(...), ...)` |
| `ExcelTabularColumn<T>` and its `ExcelDocument.WriteObjects` overload | `ExcelDocument.WriteRows(...)` |
| `CsvDocument.Map<T>(...)` | `CsvDocument.RowsAs<T>(...)` |
| `RowMapper<T>` | `OfficeIMO.Data.RowMapper<T>` |
| `CsvObjectWriter` | `CsvRowWriter` |
| `CsvObjectWriter.WriteTrustedRow(...)` / `WriteTrustedTextRow(...)` | `CsvRowWriter.WriteRow(...)` / `WriteTextRow(...)`; the established column schema still controls the overloads without a columns argument |

The generic `DbDataReader.RowsAs<T>()` extensions now live in the neutral
`OfficeIMO.Data` namespace supplied by `OfficeIMO.Core`; add
`using OfficeIMO.Data;`. CSV and Excel readers use the same mapping plan and
conversion rules. Mapping failures from this shared path throw
`DataMappingException`.

On .NET 8 and later, explicit `DateOnly` and `TimeOnly` targets work through
`RowsAs<T>`, `DbDataReader.GetFieldValue<T>`, and CSV schema columns. Default
CSV and Excel schema inference remains `DateTime`; OfficeIMO does not remap
inferred dates based on the target framework. Set
`CsvLoadOptions.MappingErrorValuePolicy` or
`ExcelReadOptions.MappingErrorValuePolicy` to `Redact` when typed mapping errors
must omit source values and custom-converter details.

Shared typed-row conversion now parses `DateTime` text with
`DateTimeStyles.RoundtripKind`. ISO 8601 round-trip values therefore preserve
their encoded UTC, local, or unspecified `DateTime.Kind` instead of being parsed
with the previous `DateTimeStyles.None` behavior. Consumers that intentionally
discard zone-kind information should normalize the mapped value explicitly.

The low-level `CsvFile` compression helper is no longer public. Use
`CsvDocument.Load`, `OpenDataReader`, `Save`, `WriteDataReader`, or caller-owned
`TextReader` / `TextWriter` streams so file and compression behavior stays with
the operation being performed.

CSV `LoadAsync` and `SaveAsync` use asynchronous source or destination I/O but
still materialize the document or serialized output. Use `OpenDataReader` for a
bounded forward-only cursor. That reader remains synchronous and can be cast to
`ICsvDataReaderPositionMetadata` for logical record numbers and available
physical start/end line numbers.

`WriteRows` keeps typed cell dispatch without boxing when its typed `Write`
overloads are used. Use `WriteRowsAsync` for an `IAsyncEnumerable<T>` source; it
consumes and disposes the source one row at a time and supports cancellation.
The async overload rejects `CreateTable` and `AutoFit` because those features
require the final row range or column widths before package output starts.

`ExcelSheet.RowsAs<T>()` is now the single typed projection name and enumerates
rows lazily while the owning workbook remains open. Replace preview
`RowsAsStream<T>()` calls with `RowsAs<T>()`; call `ToList()` or `ToArray()` when
an eagerly materialized collection is required. Use the overload accepting
`Action<RowMapper<T>>` for explicit, NativeAOT-friendly column assignments.
That mapper still requires `T : new()`. For positional records and other
constructor-bound models, use the `factory:` overload accepting
`Func<IDataRecord, T>`; it does not require a public parameterless constructor.
Named cancellation arguments on `RowsAs`, `EnumerateCells`, and `EnumerateRange`
use `cancellationToken` instead of `ct`.

Use `CsvDocument.Load(...).RowsAs<T>()` when a mutable/materialized CSV document
is required. For a forward-only typed read, use
`CsvDocument.OpenDataReader(...).RowsAs<T>()`; the same automatic and explicit
mapping definitions and the constructor factory are supported without exposing
another CSV reader type.

`ExcelDocument.Sheets` now exposes `IReadOnlyList<ExcelSheet>` instead of
`List<ExcelSheet>`. Enumerate or index the property as before, use the workbook's
worksheet operations to edit the collection, or call `document.Sheets.ToList()`
when a detached mutable snapshot is required.
Excel exposes worksheets as ordered `ExcelWorkbookDataReader` results through
`NextResult()`. Use `SheetName` or zero-based `SheetIndex` to select one sheet,
`A1Range` to select a range, `CurrentSheetName` / `CurrentSheetIndex` to identify the
workbook sheet, and `CurrentResultIndex` to identify its position in the selected results.

The Excel/CSV adapter has moved from `OfficeIMO.Reader.Excel` to the dedicated
`OfficeIMO.Excel.Csv` package and namespace. It uses the native CSV reader and
writer pipelines without making either core format package depend on the other.
Replace `ImportDelimitedFile` with `ImportCsvFile`, and
replace decoded-text `ImportDelimitedText` and worksheet `FromCsv` calls with
`ImportCsvText`. Use `ImportCsv` when the source is a `CsvDocument` or `Stream`.
Replace `ExcelDelimitedImportOptions` and `ExcelDelimitedImportResult` with
`ExcelCsvImportOptions` and `ExcelCsvImportResult`. Use `SaveAsCsv` and
`SaveAsExcel` for destination-shaped conversion entry points.

Move the removed import-option properties into the CSV parsing and reader
options owned by `ExcelCsvImportOptions`:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `Delimiter = value` | `LoadOptions.Delimiter = value` and `LoadOptions.DetectDelimiter = false` |
| `Delimiter = null` | `LoadOptions.DetectDelimiter = true` |
| `HeadersInFirstRow` | `LoadOptions.HasHeaderRow` |
| `SkipInitialRecords` | `LoadOptions.SkipInitialRecords` |
| `Culture` | `LoadOptions.Culture` |
| `ConvertNumbersAndDates` | `ReaderOptions.InferSchema` |

For example, an explicit semicolon import now uses nested options:

```csharp
using System.Globalization;
using OfficeIMO.CSV;
using OfficeIMO.Excel.Csv;

var options = new ExcelCsvImportOptions {
    SheetName = "Import",
    LoadOptions = new CsvLoadOptions {
        Delimiter = ';',
        DetectDelimiter = false,
        HasHeaderRow = true,
        SkipInitialRecords = 1,
        Culture = CultureInfo.GetCultureInfo("pl-PL")
    },
    ReaderOptions = new CsvDataReaderOptions { InferSchema = true }
};
```

`CreateTable`, `SheetName`, `TableName`, and `TableStyle` keep the same names on
`ExcelCsvImportOptions`. Leave `TableName` unset to use the effective worksheet
name, matching the former import behavior. `IncludeHeaders` controls whether the
reader's resolved field names are written into the worksheet and defaults to
`true`.

For a worksheet `FromCsv` call, move `startRow`, `startColumn`,
`firstRowIsHeader`, and `includeHeaders` into `ExcelCsvImportOptions`, and pass
`ct` as the `cancellationToken` argument. Set `CreateTable = false`, the comma
delimiter, and disabled schema inference when the migration must preserve the
former `FromCsv` no-table and string-valued parsing behavior:

```csharp
using OfficeIMO.CSV;
using OfficeIMO.Excel.Csv;

ExcelCsvImportResult imported = sheet.ImportCsvText(
    csvText,
    new ExcelCsvImportOptions {
        StartRow = startRow,
        StartColumn = startColumn,
        IncludeHeaders = includeHeaders,
        CreateTable = false,
        LoadOptions = new CsvLoadOptions {
            Delimiter = ',',
            DetectDelimiter = false,
            HasHeaderRow = firstRowIsHeader
        },
        ReaderOptions = new CsvDataReaderOptions { InferSchema = false }
    },
    cancellationToken: ct);

string range = imported.Range;
```

The former `ExecutionMode` argument has no replacement; the shared import
pipeline owns its execution strategy. The same cleanup removes `TableToCsv`;
use the worksheet `ToCsv` / `SaveAsCsv` methods instead. Calls that passed
`ToCsv` arguments positionally should use the current `headersInFirstRow`,
`csvOptions`, `readOptions`, and `cancellationToken` parameter names. Replace
old import-result properties
`RowCount`, `ColumnCount`, and `Warnings` with the current
`ExcelCsvImportResult.SheetName`, `TableName`, `Range`, and `Delimiter`.
`TableName` reports the sanitized, unique name actually created by OfficeIMO;
inspect the resulting worksheet/table when row or column counts are needed. `Delimiter`
reports the actual delimiter used, including one selected by detection.

CSV reader configuration remains in `CsvDataReaderOptions`. Excel reader safety limits remain in `ExcelReadOptions`: `MaxXlsbCells` limits aggregate workbook cells and `MaxDataReaderBufferedCells` limits a reader operation's buffer. Raise either limit only for trusted, intentionally larger workbooks.

The shared `OfficeRenderingProfile` and Excel structural mutation planning APIs are additive. Existing callers do not need compatibility wrappers for them. Use a rendering profile when multiple conversion packages must share one quality policy. Use `PlanInsertRows(...)` / `PlanDeleteRows(...)`, `PlanInsertColumns(...)` / `PlanDeleteColumns(...)`, or the range mutation plans when an application must inspect workbook impact before a transactional change; existing direct mutation calls remain available.

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

Open a source once with `PdfDocument.Open(...)`. Destination adapters also accept `PdfLogicalDocument` when an application performs custom layout analysis or page selection before conversion. Word and RTF semantic import consume shared `PdfLogicalTextRun` fragments so color, font size, and best-effort bold or italic classification do not need to be reconstructed independently in each adapter.

The common conversion grammar is:

| Intent | Shape | Example |
| --- | --- | --- |
| Return a destination model | `To{TargetModel}` | `pdf.ToWordDocument()` |
| Return a model plus diagnostics | `To{TargetModel}Result` | `pdf.ToWordDocumentResult()` |
| Return serialized content | `To{Format}` | `word.ToPdf()` |
| Write a converted artifact | `SaveAs{Format}` | `pdf.SaveAsPowerPoint(...)` |
| Write asynchronously when the operation performs asynchronous I/O | `SaveAs{Format}Async` | `pdf.SaveAsRtfAsync(...)` |
| Persist a document in its native format | `Save` / `SaveAsync` | `word.Save(...)` |
| Recover a narrow feature | Name the feature | `pdf.SaveTablesAsExcel(...)` |
| Configure forward PDF output | `{Source}PdfSaveOptions` | `WordPdfSaveOptions` |
| Configure the shared writer inside direct save options | `PdfOptions` | `HtmlPdfSaveOptions.PdfOptions` |
| Configure an intermediate conversion stage | `{Intermediate}Options` | `OneNotePdfSaveOptions.MarkdownOptions` |
| Configure reconstruction from PDF | `Pdf{Target}ImportOptions` | `PdfWordImportOptions` |
| Report a general reverse conversion | `Pdf{Target}ConversionResult` | `PdfPowerPointConversionResult` |
| Report narrow table recovery | `Pdf{Target}TableImportResult` | `PdfExcelTableImportResult` |

Target names use .NET casing: `Pdf`, `Html`, `Rtf`, `Odt`, `Ods`, `Odp`, and `PowerPoint`. Image export follows the same result-versus-write distinction: `ToImage()` opens the fluent builder, `ExportImage()` returns one structured render result, `SaveAsPng(...)` and `SaveAsJpeg(...)` write explicit encodings, and `SaveAsImages(...)` writes a page, slide, or sheet set. The public surface does not use ambiguous `SaveImage` or singular `SaveAsImage` names.

The reverse-route boundaries are:

| Route | Default result | Important limit |
| --- | --- | --- |
| PDF to Word | Semantic headings, paragraphs, lists, tables, supported images, and links | Not fixed-layout page reconstruction |
| PDF to Excel | Detected tables and structured data | Non-table page content is reported rather than placed on a worksheet canvas |
| PDF to PowerPoint | One rendered PDF page per slide | The slide image is movable, but its internals are not editable |
| PDF to PowerPoint with `EditableTables` | Detected tables on editable slides | Other page content is reported as omitted |
| PDF to RTF | Semantic text, lists, page breaks, and detected run styling | Unsupported tables, images, links, and widgets produce loss diagnostics |
| PDF to HTML | Semantic or positioned review HTML | Neither profile claims browser-clone fidelity for arbitrary PDFs |
| PDF to Markdown | Logical readable text through `pdf.Read.Markdown(...)` | Portable text rather than visual fidelity |
| PDF to ODT, ODS, or ODP | Composed Word, Excel-table, or PowerPoint-visual routes | Inspect both conversion stages and their loss reports |

`PdfResourcePolicy.CreateDefault()` is the balanced adapter default: installed and document fonts are available while arbitrary local-file and remote-resource access remains denied. Use `PdfResourcePolicy.CreatePortableDeterministic()` for untrusted or reproducible conversion. Use `PdfResourcePolicy.CreateTrustedHost()` only when the operation intentionally resolves host or remote resources.

### PowerPoint lifecycle, composition, and inspection

PowerPoint 3.1 uses the concrete `PowerPointPresentation`, `PowerPointSlide`, and shape types as the editing model. Semantic deck plans, template helpers, and format adapters remain optional workflows over that model.

The `OfficeIMO.PowerPoint.Fluent` namespace and its PowerPoint-only builders, the old designer extensions, and the public `PowerPointDeckComposer` were removed. Replace fluent concrete-editing calls with `PowerPointPresentation`, `PowerPointSlide`, and concrete `PowerPointShape` operations. Replace semantic composer or designer calls with a `PowerPointDeckPlan` followed by `PowerPointPresentation.Compose(...)`. Custom semantic slides remain available through `PowerPointDeckPlan.AddCustom(...)`; its callback receives a `PowerPointSlideCompositionContext`.

Replace lifecycle calls as follows:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `Open(path)` | `Load(path)` |
| `OpenRead(path)` | `Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })` |
| `Open(stream, readOnly: true, autoSave: false)` | `Load(stream, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })` |
| `Create(stream, autoSave: false)` | `Create(stream)` |
| Implicit save on dispose | Set `PersistenceMode = DocumentPersistenceMode.SaveOnDispose`, or call `Save()` explicitly |

`PowerPointPresentation.Create(...)` now starts with zero slides, and every `AddSlide()` call creates one new slide. Replace code that indexed or reused the old placeholder at `Slides[0]` with an explicit `AddSlide()` before accessing the new slide.

The designer and deck-composer entry points now route through one plan and one composition call:

| OfficeIMO 3.0 | OfficeIMO 3.1 |
| --- | --- |
| `OfficeIMO.PowerPoint.Fluent` concrete-editing builders | Use `PowerPointPresentation`, `PowerPointSlide`, and concrete `PowerPointShape` operations directly |
| `PowerPointDeckComposer.Add...` or other public composer calls | Add the semantic slide to `PowerPointDeckPlan`, then call `PowerPointPresentation.Compose(...)` |
| `AddDesigner...` extension methods | Use the corresponding `PowerPointDeckPlan.Add...` method, then call `Compose(...)` |
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

For table-only recovery, `HasLoss` means a detected table was truncated by an import limit. `HasOmittedPageContent` means the source also contains non-table text, vectors, images, links, forms, annotations, or actions that the adapter does not import. Use `SourceScope` for the counts behind that decision. Choose Word or RTF semantic conversion, or a rendered-page route, when the goal is a broader page representation.

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

`LegacyXlsLoadResult.AdvancedWorkbook` is the public imported workbook. Replace `LegacyXlsLoadResult.CreateAdvancedImportReport()` and the old `ImportReport` property with the cached `CreateImportReport()` result.

The `OfficeIMO.Core` target-framework compatibility type `System.Runtime.CompilerServices.IsExternalInit` is internal in the `netstandard2.0` and `net472` assets. Remove any application reference to that shim; normal record and `init` usage remains supported.

Markdown-to-Word callers should parse through `OfficeIMO.Word.Markdown` rather than calling the removed inline-run helper directly. `ConvertDotxToDocx(...)` also resolves relative template paths before package URI construction, so relative and absolute template paths use the same behavior.

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

### Shared foundation package

The compiled `OfficeIMO.Shared` implementation package no longer exists. `OfficeIMO.SharedSource` is source-only and is not a runtime package replacement. Move direct package references to the public owner of each reusable value: shared colors, fonts, images, charts, lifecycle options, stream contracts, export results, and dependency-free security provider contracts belong to `OfficeIMO.Core`; normalized Reader contracts belong to `OfficeIMO.Reader.Core`; the optional concrete CMS, XML DSig, X.509, and RFC 3161 provider belongs to `OfficeIMO.Security`. Drawing APIs remain in the `OfficeIMO.Drawing` namespace, lifecycle APIs use the root `OfficeIMO` namespace, and native document behavior remains in its format package.

OfficeIMO 3.1 introduces the zero-dependency `OfficeIMO.Core` package by renaming the former `OfficeIMO.Drawing` package, project, and assembly; actual drawing APIs remain in the `OfficeIMO.Drawing` namespace. Native packages still own parsing, loading, editing, validation, and serialization for their formats. Adapter packages project one native model into another rather than exposing another parser or document model. `OfficeIMO.Html` owns the canonical HTML source model and resource policy; format adapters consume it. These ownership changes replace direct use of the former shared implementation layer without introducing another dependency tier.

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

The canonical forward PDF result method is `ToPdfDocumentResult()`. Reverse PDF adapters extend `PdfDocument` and `PdfLogicalDocument` with destination-shaped result methods such as `ToWordDocumentResult()`, `ToPowerPointPresentationResult()`, and `ToRtfDocumentResult()`. `SaveAsPdf(...)` returns `PdfSaveResult` evidence across Word, Excel, PowerPoint, HTML, Markdown, and RTF adapters, while `ToPdf()` remains the encoded-byte convenience API. Opening a generated file in another application is an explicit application action, not part of saving.

`VisioDocument.Load(path)` and `Load(stream)` now apply a 512 MiB default input
limit before opening the package. For trusted documents that intentionally
exceed that size, pass `new VisioLoadOptions { MaxInputBytes = null }`. The
options-first async overload keeps cancellation explicit:
`LoadAsync(path, options, cancellationToken)`. The token-first overloads were
removed. Replace `LoadAsync(path, cancellationToken)` with
`LoadAsync(path, cancellationToken: cancellationToken)`, and replace
`LoadAsync(path, cancellationToken, options)` with
`LoadAsync(path, options, cancellationToken)`. Use the same ordering for the
stream overload.

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
| `SavePdfTablesAsExcel/Word/PowerPoint()` | `SaveTablesAsExcel()` / `SaveAsWord()` / `SaveAsPowerPoint()` |
| `ToPngResult` / `ToSvgResult` and plural result aliases | `ExportImage()` / `ExportImages()` returning `OfficeImageExportResult` values |
| `PdfImageExportOptions.MaxPages` | `MaximumOutputCount` or `ToImages().WithMaximumPages(...)` |
| `EmailDocument.WriteToBytes()` | `EmailDocument.ToBytes()` |

Format-spelling aliases such as `SaveToPdf`, `SaveAsBytesToPdf`, and generic `WriteToBytes` were removed. Use `SaveAsPdf(...)` for a destination and `ToPdf()` or `ToBytes()` for an in-memory result. Ambiguous `SaveImage` / `SaveAsImage` names were replaced by explicit encodings such as `SaveAsPng(...)`, or by `SaveAsImages(...)` for multi-page and multi-sheet output.

Image export uses `OfficeImageExportResult` and `OfficeImageExportFormat` from the `OfficeIMO.Drawing` namespace supplied by `OfficeIMO.Core`. Replace the removed scale presets as follows:

| Removed member | Replacement |
| --- | --- |
| `WithDpi(...)` | `AtDpi(...)` for physical output density |
| `ForHighResolution(...)` | `ForPrint(...)` for the print profile |

The byte-returning helpers remain encoding-specific: `ToPng()`, `ToJpeg()`, `ToTiff()`, and `ToWebp()` return bytes, while `ToSvg()` returns SVG text. Format-specific saves such as `SaveAsPng(...)`, `SaveAsJpeg(...)`, and the fluent `As...().Save(...)` surface write a destination and return structured evidence. `WithScale(...)` remains available for renderer-relative scaling.

Image file saves now default to `OfficeImageExportFileConflictPolicy.FailIfExists`. A repeated write to the same path therefore throws unless the caller explicitly selects `Replace` or `CreateUnique` with `OnFileConflict(...)`.

Raster exports share a 50-million-output-pixel default. The default overflow policy reduces scale before allocating the pixel buffer and reports `IMAGE_RASTER_SCALE_REDUCED`; set `RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw` to receive an `OfficeImageExportLimitException` instead.

`OfficeImageExportResult` validates that encoded bytes, format, and dimensions agree. `DpiX`, `DpiY`, `PhysicalWidthInches`, `PhysicalHeightInches`, and `EncodedLength` are derived from the encoded payload. Shared options own `MaximumRasterPixels`, `RasterOverflowBehavior`, `ImageCodec`, `RasterEncoding`, `TargetDpi`, `Fonts`, `Policy`, `Progress`, batch limits, and maximum concurrency; document-specific options inherit those values rather than redefining them.

Batch builders use `ExportEach(...)` / `ExportEachAsync(...)` for bounded streaming and `SaveFiles(...)` / `SaveFilesAsync(...)` when callers need saved path, metadata, and diagnostics without retaining every encoded payload. `OfficeImageExportPolicy` can reject loss, omissions, failures, or selected diagnostic codes. Supply intended TrueType faces through `WithFont(...)`, `WithFonts(...)`, or `OfficeImageExportOptions.Fonts` when `OfficeImageExportDiagnosticCodes.FontSubstituted` would otherwise be reported.

PDF image export uses the same builder. Use `PdfReadPage.ToDrawing()` when an application needs the intermediate `OfficeDrawing` scene. `PdfPageRenderResult` remains a lower-level inspection, OCR, and verification contract with timing and PDF capability diagnostics; it is not the general multi-format export result.

Format-neutral SVG image export now writes whole-pixel `px` root dimensions so the encoded size matches `OfficeImageExportResult.Width` and `Height`. If a lower-level Drawing workflow requires physical point units, call `OfficeDrawingSvgExporter.ToSvg(...)` with `OfficeSvgSizeUnit.Point` explicitly.

Image decode and font-fallback diagnostics moved to the shared Drawing result:

| Removed diagnostic | Replacement |
| --- | --- |
| `ExcelImageRasterFormatUnsupported`, `ExcelImageSvgFormatUnsupported`, `ExcelImagePngDecodeUnavailable`, `ExcelHeaderFooterImageUnsupported` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `unsupported-word-image-raster` / `unsupported-word-image-svg` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `unsupported-powerpoint-image-raster` / `unsupported-powerpoint-image-svg` | `IMAGE_SOURCE_DECODE_FALLBACK` |
| `HtmlRenderRasterDecoderUnavailable` | `IMAGE_SOURCE_DECODE_FALLBACK` on the final image export result |
| `ExcelCellFontFamilyFallback`, `ExcelChartFontFamilyFallback`, `ExcelHeaderFooterFontFamilyFallback` | `IMAGE_FONT_SUBSTITUTED` |

`IMAGE_SOURCE_DECODED_BY_CALLER_CODEC` records that a caller-supplied `ImageCodec` handled the source.

### HTML, Reader, and theme ownership

Raw HTML is parsed once into `HtmlConversionDocument`. PDF, image, Word, Markdown, RTF, Excel, and PowerPoint adapters consume that model so the caller's base URI, document `<base>` semantics, source DOM, URL policy, media intent, and diagnostics are not reinterpreted by separate parsers. Replace adapter-specific raw-HTML parsing with `HtmlConversionDocument.Parse(...)`, then call the destination adapter.

Reader orchestration uses an immutable `OfficeDocumentReader` built from explicit typed handlers. Native packages retain parser ownership; Reader adapters project native models into `OfficeDocumentReadResult` and return diagnostics from the operation. Do not replace removed parser classes with another public Reader parser hierarchy.

Markdown HTML and PDF options use the shared `MarkdownVisualTheme` through `Theme`; PDF-only overrides use `MarkdownPdfSaveOptions.Style` and `MarkdownPdfStyle.DocumentTheme`. Visio styling and package themes are separate contracts: `VisioStyleTheme` describes reusable diagram styling, while `VisioPackageTheme` represents theme data stored in a Visio package. Shared colors and hexadecimal formatting belong to `OfficeIMO.Drawing` rather than duplicate Word or Excel helpers.

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
