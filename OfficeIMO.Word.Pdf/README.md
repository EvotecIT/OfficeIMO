# OfficeIMO.Word.Pdf - Word/PDF conversion

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.Pdf)](https://www.nuget.org/packages/OfficeIMO.Word.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.Pdf)

`OfficeIMO.Word.Pdf` exports `OfficeIMO.Word` documents to PDF through the first-party `OfficeIMO.Pdf` engine and imports parser-supported PDF logical content into editable Word documents. It is the adapter layer: Word stays responsible for the `.docx` model, while PDF layout, reading, diagnostics, and writing stay in `OfficeIMO.Pdf`.

## Legacy DOC to PDF

`LegacyDocPdfConverter` composes bounded binary DOC import with the Word PDF adapter. It blocks known import loss by default and retains the import report separately from PDF rendering diagnostics:

```csharp
using OfficeIMO.Word.Pdf;

using var source = File.OpenRead("archive.doc");
var conversion = LegacyDocPdfConverter.ToPdfDocumentResult(source);
conversion.SaveResult("archive.pdf").RequireSuccess();
foreach (var report in conversion.SourceConversionReports)
    foreach (var finding in report.FidelityDiagnostics)
        Console.WriteLine($"{finding.Code}: {finding.Message}");
```

Pass `lossPolicy: OfficeIMO.OfficeConversionLossPolicy.Allow` only when accepting the reported import reductions. Import errors still block output. The adapter always collects unsupported-content findings, even if the supplied import options disable reporting. Its fidelity is limited by both the legacy importer and the Word PDF renderer; it does not guarantee exact Microsoft Word pagination or rendering of every binary DOC feature.

## Install

```powershell
dotnet add package OfficeIMO.Word.Pdf
```

## Quick start

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Create("report.docx");
document.AddParagraph("PDF export").Style = WordParagraphStyles.Heading1;
document.AddParagraph("This document is exported through OfficeIMO.Pdf.");

document.SaveAsPdf("report.pdf");
```

## Examples

### Export with page and metadata options

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("proposal.docx");

var options = new WordToPdfOptions {
    Orientation = OfficePageOrientation.Portrait,
    Margins = PageMargins.UniformCentimeters(1.5),
    Title = "Customer proposal",
    Author = "Evotec",
    IncludePageNumbers = true,
    PageNumberFormat = "Page {current} of {total}"
};

document.SaveAsPdf("proposal.pdf", options);
```

### Refresh date fields before export

PDF export uses the field results stored in the Word document by default. To refresh dynamic `DATE` and `TIME` fields for the PDF, set `DateTimeFieldUpdateOptions`:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("report.docx");
document.SaveAsPdf("report.pdf", new WordToPdfOptions {
    DateTimeFieldUpdateOptions = new WordFieldUpdateOptions()
});
```

`DATE` and `TIME` use the current local clock during the refresh. Set `CurrentDateTime` on `WordFieldUpdateOptions` when the export needs a fixed value. The refresh updates the in-memory Word document; it does not save the `.docx` file or change other fields. `CREATEDATE` and `SAVEDATE` continue to use their cached results unless explicitly refreshed through the general Word field API.

Table borders follow the Word style and direct cell settings: `nil` suppresses a shared edge while `none` yields to the opposing border. Set `DefaultTableBorders = true` only when you want a fallback grid on otherwise borderless tables.

Inline pictures in table cells retain their position among the paragraph's text, including multiple pictures in one run. Their authored display dimensions and alternative descriptions pass to the shared PDF layout; exact-height lines can clip them. Mixed picture runs retain visible text, and hidden runs suppress their pictures in body content, tables, headers and footers. Nested-table content is flattened, so preserving its pictures does not preserve nested table frames.

Ordinary underlining includes spaces between words. Word's explicit *underline words only* style continues to leave those spaces clear.

Font sizes preserve half-point values, including 10.5 pt. When the run and its styles omit a size, conversion honors the document default. An existing `docDefaults` element without a size uses Word's 10 pt fallback; a document without `docDefaults` uses 12 pt. OfficeIMO-created documents declare an 11 pt default and retain that size.

### Preserve fonts from a trusted template

Balanced conversion substitutes standard PDF fonts for document-selected fonts. This can change line breaks, table text width, and baselines. For a trusted template whose fonts are installed on the host, enable document-font embedding:

```csharp
var options = new WordToPdfOptions();
options.ResourcePolicy.AllowDocumentFontEmbedding = true;
document.SaveAsPdf("template.pdf", options).RequireSuccess();
```

Both system-font and document-font embedding must be allowed. This setting retains the template's individual font families; `FontFamily` instead selects a conversion-wide default. Substitution warnings identify when the resource policy disables embedding, separately from an unavailable font. Local-file and remote-resource access remain governed by their own policy settings.

Positioned tables in ordinary document flow preserve page, margin, or text anchors, explicit offsets, and text clearances. Following paragraphs use the available space beside the table and return to full width below it. Headings, lists, images, and other structured blocks move below an intersecting table. Positioned tables in multi-column sections retain an approximation warning.

Section columns use the shared PDF flow engine for equal and unequal widths, continuous section transitions and final-page balancing. Paragraph keep, widow and spacing settings remain active inside the column frame. Tables split at row or supported cell-content boundaries, and repeated headers stay with their first body row. Signed table indentation is retained during DOCX/native DOC authoring and PDF placement; an explicit zero indent overrides inherited table-style indentation.

Tables with positive cell spacing retain an outer perimeter and separate cell borders, including table shading in the gaps and different border colors, widths and supported patterns. The frame follows table alignment, indentation, merged cells, page and column continuation, and repeating headers. Exact row heights retain partially visible text through cell clipping; minimum heights allow content to grow. Set `table.StyleDetails.CellSpacing` in twips to author the spacing. An imported explicit automatic or percentage spacing value clears inherited spacing. Floating tables and nested table layout remain subject to the native engine's supported paths.

Paragraph border spacing positions the border outside the text frame without adding horizontal text padding. Borders spanning columns or pages keep their side strokes; their top belongs to the first fragment and their bottom to the last. Native DOC compatibility can produce different fragment capacities from modern DOCX. These mappings preserve the supported source settings; font substitution and unsupported layout features can still change pagination.

Line spacing follows document defaults, table styles, paragraph styles and direct formatting, including built-in headings. Automatic spacing uses the effective paragraph and run fonts after substitution or embedding; exact and minimum spacing retain their point units. Rich body and table paragraphs use their own font size during measurement and pagination. A large run on another line or a large paragraph mark does not impose a minimum font size on every rich line. Authored line breaks retain their run formatting, so larger blank lines can expand minimum spacing. Exact spacing keeps a fixed advance; minimum spacing can expand for larger runs or inline elements. An authored line value without a rule uses automatic spacing. A rule without a numeric line value inherits the complete spacing pair. Font substitution can still change line advances. First-baseline placement, baseline offsets between mixed-size lines, clipping within exact-height lines and exact Word pagination remain limited.

Empty paragraphs retain a blank line when their paragraph mark is visible, including paragraphs whose text runs are all hidden. Their line metrics use the paragraph mark and inherited typography. A hidden mark removes the empty line and its paragraph spacing; an explicit visible mark overrides a hidden paragraph style. An image, shape or inline group rendered in paragraph flow reserves its object height once; a following empty paragraph retains its own blank line. Adjacent text paragraphs in the body, columns and block content controls join across hidden paragraph marks. Joined text uses the first paragraph's alignment and spacing before, the final paragraph's spacing after, and each source run's own inherited typography. Lists, headings, decorated paragraphs, explicit page or column breaks, and object paragraphs retain separate boundaries and emit a warning. Joining inside table cells, headers and footers remains limited.

Object-only paragraphs retain direct and inherited spacing before and after their flow objects, collapsing that spacing with adjacent paragraphs. Hidden shapes with a resolved flow height reserve that height without drawing the shape. Positioned-shape placement remains subject to the renderer's supported anchor and wrapping mappings.

### Export to bytes or streams

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("invoice.docx");

byte[] pdfBytes = document.ToPdfBytes();

using var stream = File.Create("invoice.pdf");
document.SaveAsPdf(stream);
```

### Optimize images during PDF export

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("input.docx");
byte[] pdf = document.ToPdfBytes(new WordToPdfOptions {
    PdfOptions = new PdfOptions {
        ImageOptimization = new PdfImageOptimizationOptions {
            Enabled = true,
            Mode = OfficeImageOptimizationMode.DownsampleAndRecompress,
            TargetDpi = 144,
            JpegQuality = 85
        }
    }
});
File.WriteAllBytes("output.pdf", pdf);
```

PDF image optimization is opt-in and leaves Word source media unchanged. `Downsample` uses the final source-image placement, including crop/fit expansion; `Recompress` retains pixels and re-encodes JPEGs; `DownsampleAndRecompress` applies both. The shared managed codecs handle static PNG/JPEG/BMP/GIF/TIFF/WebP input. Multi-frame/page payloads and unsupported formats remain outside static optimization. Candidates that would grow or remove metadata are preserved by default; `AllowMetadataLoss` permits reported loss or intentional stripping. Explicit metadata policies also apply when no pixel reduction is needed. The PDF layout report records optimization decisions and warns when metadata is removed. This export policy is separate from the lossless optimizer for existing PDFs.

### Capture conversion warnings without throwing away the report

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("complex-document.docx");
var options = new WordToPdfOptions {
    DefaultTableBorders = true
};

var result = document.SaveAsPdfResult("complex-document.pdf", options);
if (!result.Succeeded) {
    foreach (string diagnostic in result.Diagnostics) {
        Console.WriteLine(diagnostic);
    }
}

foreach (var warning in result.Warnings) {
    Console.WriteLine($"{warning.Source}: {warning.Message}");
}
```

### Import semantic PDF content into Word

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

PdfDocument pdf = PdfDocument.Load("packet.pdf");
PdfWordConversionResult import = pdf.ToWordDocumentResult(
    new PdfToWordOptions());

using WordDocument word = import.Value;
word.Save("packet.docx");

foreach (var warning in import.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The semantic import path preserves document metadata, page breaks, headings, paragraphs, lists, detected run color/size/emphasis, logical tables, safe URI hyperlinks, supported internal destination links, embedded images, form-widget placeholders, and conversion diagnostics when those structures are available in the PDF logical model. `UseSharedPageReadingOrder` defaults to `true`, so semantic items follow the core engine's crop-, rotation-, spanning-band-, and column-aware order. Editable output also retains each source page's physical size, uses a 36-point page margin, suppresses extra paragraph spacing in dense source content, preserves detected table-column proportions, and positions supported axis-aligned images relative to the page. These defaults improve business-document reconstruction, but they do not turn semantic import into fixed-layout page recreation or Microsoft Word rendering parity.

Set `PreserveSourcePageSize`, `PreserveCompactSourceSpacing`, or `PreserveImagePlacementPosition` to `false` when normal Word flow is more useful than the corresponding source geometry. `EditablePageMarginPoints` controls the source-sized section margin. Source page sizes require `PreservePageBreaks = true` when more than one PDF page is imported.

For selected pages, Fast versus Structured reconstruction, or custom semantic budgets, pass the canonical read settings through the import options:

```csharp
PdfWordConversionReport report = pdf.SaveAsWord(
    "packet-selection.docx",
    new PdfToWordOptions {
        ReadOptions = new PdfReadOptions {
            PageSelection = PdfPageSelection.Parse("1-3,5"),
            LayoutOptions = new PdfTextLayoutOptions { ForceSingleColumn = true },
            Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 5 }
        }
    }).RequireSuccess().Report!;
```

`ReadOptions` is used only when the source is an opened `PdfDocument`. When a `PdfDocumentReadResult` is already available, the adapter converts that supplied logical model directly.

For table-only recovery, use the same façade with the explicit profile:

```csharp
pdf.SaveAsWord(
    "statement-tables.docx",
    PdfToWordOptions.CreateTablesOnly());
```

## What it exports

- Paragraphs, headings, rich runs, links, bookmarks, page breaks, lists, and common spacing/indentation settings, including hanging and legal negative left/right indents. Built-in Word heading levels 1–9 use the same level mapping as the table of contents in body and column flow. Set `WordToPdfOptions.PdfOptions.TaggedStructureMode` to `PdfTaggedStructureMode.CatalogMarkers` to retain explicit numeric levels during PDF reading and editable Word import, including skipped levels. Tagged PDF uses standard-compatible role mappings for levels 7–9. Untagged output retains the bookmark hierarchy, whose nesting depth cannot identify skipped numeric levels.
- Word-authored text bullets use portable marker characters. Picture bullets currently use a text bullet in PDF output and report `NativePictureBulletTextFallback` with the source picture-bullet identifier; the embedded marker image is not rendered.
- Word sections, page size, orientation, margins, columns, headers, footers, page numbers, and document background color.
- Tables with common Word table styling, repeated headers, cell fills, borders, alignment, merged cells, and rich text in cells.
- Paragraph-aligned images, selected shapes, text boxes, content controls, simple form controls, footnote/endnote markers, and table-of-contents links where supported by the first-party PDF path.
- DrawingML groups of supported preset shapes retain nested child coordinates and scaling. Non-wrapping groups behind text preserve page- or margin-relative positions and paragraph-relative vertical anchors. Paragraph anchors follow pagination, columns, and floating-table clearance; list paragraphs retain both their marker and the group.

- Unrotated, uncropped `InFrontOfText` images with explicit page-relative offsets inside the page bounds. Images follow the first page of their anchor paragraph or heading, including section columns, and paint over text and other flow content without reserving their height in the document flow. Overlapping foreground images follow `WordImage.ZOrder`; images with equal values retain document order.
- Per-operation conversion warnings through `PdfDocumentConversionResult.Report` or `PdfSaveResult.Report`.

Section gutters reserve space at the left, right, or top of the body frame according to the document settings. Mirrored margins swap the left and right body margins on even visible page numbers, including section numbering restarts. Margin-relative shape groups follow that frame; page-relative groups retain their absolute coordinates. A top gutter uses the same horizontal margins on both page sides, matching Word. An explicit `WordToPdfOptions.Margins` replaces the authored margins, gutter, and mirroring.

For imported groups with unsupported DrawingML geometry, fixed-position export uses the document's VML fallback when available and reports `NativeShapeGroupVmlFallback`. Supported groups with other wrapping or anchor modes are placed in document flow with `NativeShapeGroupFlowed`; groups that cannot be rendered report `NativeShapeGroupUnsupported`. Arbitrary custom geometry, rotation, flips, foreground stacking, and exact text wrapping around groups remain limited.

Next-page section starts create a new page. Odd/even starts use the continuing page number to insert a blank page when needed, then apply the new section's numbering restart. An odd/even start advances a conflicting restart to the next matching number. Next-page starts with an explicit restart also align the section start with that number's parity. Compatible continuous sections share a page. Embedded page breaks preserve text, run formatting, hyperlinks and explicit bookmark targets on both sides of the break, including consecutive blank pages. Fields spanning paragraphs retain their result visibility through page and column splits; hidden field instructions and their breaks do not create pages. A section mark without body content retains its editable formatting and anchors without adding a blank body line or an extra page. Column layout and changes in page geometry remain subject to the native layout engine's supported paths.

## What it imports

- Parser-supported PDF metadata, page breaks, headings, paragraphs, lists, logical tables, safe URI hyperlinks, supported internal destination links, complete image-file payloads with transparency-mask fidelity metadata, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware soft-mask-capable simple `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware soft-mask-capable `Indexed` palette PDF image streams into editable `.docx` content when their filters are supported.
- Image fallback placeholders and form-widget placeholders with diagnostics instead of silently dropping unsupported objects.
- Page-range filtered imports through `PdfDocument.Read(new PdfReadOptions { PageSelection = ... })`.
- Active hyperlink reconstruction for absolute `http`, `https`, and `mailto` URI annotations through `PdfToWordOptions.ImportUriLinks` and `PdfToWordOptions.AllowedHyperlinkUriSchemes`.
- Internal PDF destination reconstruction through `PdfToWordOptions.ImportInternalLinks`, mapping supported page and named destinations to Word bookmarks and anchor hyperlinks.
- Native image embedding through `PdfToWordOptions.ImportImages`; complete image files, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware soft-mask-capable simple 8-bit `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware soft-mask-capable `Indexed` palette streams are embedded when their filters are supported. Fully transparent and unplaced resources are suppressed. Images whose clips, placement soft masks, unresolved source transparency masks, or unsupported blend modes could expose hidden source pixels are not embedded as raw pictures; the report records a typed omission, and `PdfToWordOptions.IncludeImagePlaceholders` can retain an editable marker instead. Supported opacity is mapped to native Word picture transparency.
- Per-operation import warnings through `PdfWordConversionResult.Report`. `HasLoss` and `RequireNoLoss()` use typed approximation, omission, and failure evidence even when a diagnostic is informational. The table-only profile reports visible text outside detected tables as `PdfTextContentNotImported` instead of treating an intentionally narrow extraction as lossless.

## Options and diagnostics

For a PDF whose page appearance matters more than editability, use visual pages:

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

var pdf = PdfDocument.Load("source.pdf");
var options = PdfToWordOptions.CreateVisualPages();
options.Dpi = 144;
options.ReadOptions = new PdfReadOptions {
    PageSelection = PdfPageSelection.Parse("1-3,5")
};
pdf.SaveAsWord("visual-pages.docx", options);
```

Each selected page becomes an image on a Word section with the source page's physical dimensions. Text, links, and form fields are not editable in this mode. Rendering uses the managed PDF engine and reports its limitations. The default limits are 100 pages, 64 million pixels per page, and 256 MB of encoded page images; Word pages larger than 22 inches in either dimension are rejected. Use the default `EditableContent` mode to reconstruct supported text, tables, and images as Word objects.

Use `WordToPdfOptions` when callers need to override page geometry, metadata, page-number behavior, font family, table-border fallback, profile presets, or text fallback policy. `TextFallbacks` uses the shared `PdfTextFallbackFeatures` enum. The balanced resource default enables installed fonts but denies arbitrary local and remote reads; use `PdfResourcePolicy.CreatePortableDeterministic()` for reproducible or untrusted conversion and `CreateTrustedHost()` only when local or remote resource access is intentional. Profiles do not inject page numbers; set `IncludePageNumbers = true` explicitly when generated numbering is desired. Request `ToPdfDocumentResult()` or `SaveAsPdfResult()` when diagnostics matter; unsupported Word features and preserved header/footer overflow become actionable operation results instead of mutable option state. Available embeddable Word families use shared named PDF resources and are not limited to three compatibility slots. Unavailable or non-embeddable families fall back to a mapped PDF font with an explicit warning.

## Current limits

- Floating images outside the supported page-relative placement contract use document flow and report `NativeAnchoredImageFlowed`. Complex wrapping and overlapping-object layout are not reconstructed.
- This package does not try to be a full Word renderer with perfect Microsoft Word parity or a fixed-layout PDF-to-DOCX recreation engine.
- Editable PDF-to-Word import reconstructs parser-supported logical objects. Complex or unsupported PDF image streams, interactive controls, unresolved destinations, and remote or cross-document navigation actions are not reconstructed as native Word objects. Visual pages preserve rendered appearance as images within the managed renderer's support and configured limits. Open reconstruction work is tracked in the repository [roadmap](../Docs/ROADMAP.md).

Use `OfficeIMO.Pdf` for direct PDF layout and manipulation. PowerShell workflows are available through [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice).

## Related packages

- [OfficeIMO.Word](../OfficeIMO.Word/README.md) - Word document model.
- [OfficeIMO.Pdf](../OfficeIMO.Pdf/README.md) - PDF creation, reading, and manipulation engine.
- [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) - HTML/PDF bridge built on OfficeIMO converters.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages; no browser, native renderer, or commercial PDF SDK.
- **OfficeIMO:** `OfficeIMO.Word`, `OfficeIMO.Pdf`, and `OfficeIMO.Core` own the source model, PDF engine, mapping, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 1 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
