# OfficeIMO 4.0 PDF bridge migration

OfficeIMO 4.0 gives each optional PDF adapter one discoverable surface in both directions. Open a PDF once with `PdfDocument.Open(...)`, then call the destination-shaped method supplied by the package you installed.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

PdfDocument pdf = PdfDocument.Open("source.pdf");
PdfWordConversionResult result = pdf.ToWordDocumentResult();

using OfficeIMO.Word.WordDocument word = result.Value;
word.Save("source.docx");
```

The same pattern applies to Excel, PowerPoint, HTML, and RTF:

```csharp
pdf.SaveAsExcel("tables.xlsx");
pdf.SaveAsPowerPoint("pages.pptx");
pdf.SaveAsHtml("review.html");
pdf.SaveAsRtf("editable.rtf");
```

Every editable destination also accepts an already loaded `PdfLogicalDocument`. Use that lower-level receiver when you need custom layout analysis or page selection:

```csharp
PdfLogicalDocument selected = pdf.Read.Logical(
    PdfPageSelection.Parse("1-3,5"),
    new PdfTextLayoutOptions { ForceSingleColumn = true });

selected.SaveAsWord("selected.docx");
```

## OpenDocument packages

The 3.x `OfficeIMO.OpenDocument.Pdf` package pulled Word, Excel, and PowerPoint adapters together. In 4.0 it is replaced by focused packages so applications carry only the route they use:

| Route | Focused package | Reverse entry point |
| --- | --- | --- |
| ODT ⇄ PDF | `OfficeIMO.OpenDocument.Odt.Pdf` | `pdf.ToOdtDocument()` |
| ODS ⇄ PDF | `OfficeIMO.OpenDocument.Ods.Pdf` | `pdf.ToOdsDocument()` |
| ODP ⇄ PDF | `OfficeIMO.OpenDocument.Odp.Pdf` | `pdf.ToOdpPresentation()` |

There is no umbrella or bridge-specific Core package in 4.0. Install the format adapter your application actually uses.

Each reverse result exposes the native PDF import report and the OpenDocument feature-mapping report. PDF-to-ODS reports non-table page content as loss. PDF-to-ODP defaults to visual pages when the receiver is an opened `PdfDocument`; the lower-level logical receiver supports the editable-table profile because visual rendering needs the original PDF bytes.

## API renames

The 4.0 boundary removes the overlapping 3.x names instead of retaining aliases.

| OfficeIMO 3.x | OfficeIMO 4.0 |
| --- | --- |
| `PdfSaveOptions` in `OfficeIMO.Word.Pdf` | `WordPdfSaveOptions` |
| `PdfWordReadOptions` | `PdfWordImportOptions` |
| `PdfRtfReadOptions` | `PdfRtfImportOptions` |
| `PdfExcelTableImportOptions` | `PdfExcelImportOptions` |
| `PdfExcelTableImportReport` / `Result` | `PdfExcelImportReport` / `Result` |
| `ImportTablesToExcelDocument` | `ToExcelDocument` |
| `SaveTablesAsExcel` | `SaveAsExcel` |
| `PdfPowerPointTableImportOptions` | `PdfPowerPointImportOptions` |
| `PdfPowerPointTableImportReport` / `Result` | `PdfPowerPointImportReport` / `Result` |
| `ImportTablesToPowerPointPresentation` | `ToPowerPointPresentation` |
| `SaveTablesAsPowerPoint` | `SaveAsPowerPoint` |

`PdfWordImportOptions.CreateTablesOnly()` and `PdfPowerPointImportOptions.CreateEditableTables()` replace separate table-only façades when that narrower result is intentional.

## What each reverse route produces

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

Word and RTF semantic import now consume shared `PdfLogicalTextRun` fragments. Those fragments preserve detected source color, font size, and best-effort bold/italic classification without making each destination adapter realign raw PDF spans independently.

## PowerPoint modes

The opened-PDF PowerPoint route defaults to visual pages because a page image is a more useful and honest general PDF-to-slide result than returning only detected tables:

```csharp
PdfDocument pdf = PdfDocument.Open("handout.pdf");
PdfPowerPointImportReport report = pdf.SaveAsPowerPoint("handout.pptx");
```

For editable table recovery:

```csharp
var options = PdfPowerPointImportOptions.CreateEditableTables();
options.MaxRowsPerSlide = 18;
options.MaxColumnsPerSlide = 6;

PdfPowerPointImportReport report = pdf.SaveAsPowerPoint(
    "handout-tables.pptx",
    options);
```

The current visual mode is the foundation for a later hybrid mode with editable text and image layers. Arbitrary PDF vectors, groups, clipping, forms, annotations, and presentation animations are not claimed as editable PowerPoint objects.

## Resource defaults

`PdfResourcePolicy.CreateDefault()` is the balanced fidelity default for PDF adapter packages. It permits installed-font and document-font embedding while continuing to deny arbitrary local-file and remote-resource access.

Use `PdfResourcePolicy.CreatePortableDeterministic()` for untrusted or reproducible jobs that must not inspect host fonts. Use `CreateTrustedHost()` only when a conversion intentionally resolves local or remote resources.

Word-to-HTML now emits detected run colors and highlights by default. Set `IncludeRunColorStyles` or `IncludeRunHighlightStyles` to `false` only when a deliberately style-reduced HTML result is required.

## Roadmap gaps

Reverse routes remain supported and should expand where the destination model can represent useful content:

- PDF → Excel: improve table continuation, repeated-header recognition, typed values, and bounded positioned-cell recovery; do not present arbitrary page art as a workbook.
- PDF → PowerPoint: add a hybrid visual/editable mode, then reconstruct bounded text boxes and supported image layers while retaining the rendered page as an optional reference.
- PDF → Word and RTF: extend shared run reconstruction, table/image coverage, and positioning diagnostics before attempting broad page-layout claims.
- PDF → HTML: keep semantic and positioned profiles explicit, and improve shared asset/style diagnostics rather than merging them into an ambiguous default.

Each expansion needs an artifact test and a truthful report for content that remains simplified or omitted.
