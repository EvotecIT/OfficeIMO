# OfficeIMO.Word.Pdf - Word/PDF conversion

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.Pdf)](https://www.nuget.org/packages/OfficeIMO.Word.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.Pdf)

`OfficeIMO.Word.Pdf` exports `OfficeIMO.Word` documents to PDF through the first-party `OfficeIMO.Pdf` engine and imports parser-supported PDF logical content into editable Word documents. It is the adapter layer: Word stays responsible for the `.docx` model, while PDF layout, reading, diagnostics, and writing stay in `OfficeIMO.Pdf`.

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

PDF export uses the field results stored in the Word document. To calculate supported fields such as `DATE`, `CREATEDATE`, and `SAVEDATE` first, refresh them explicitly:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("report.docx");
document.UpdateFieldsAndGetReport();
document.SaveAsPdf("report.pdf");
```

`DATE` uses the current clock during the refresh. `CREATEDATE` and `SAVEDATE` use the document's stored properties. Table borders follow the Word style and direct cell settings: `nil` suppresses a shared edge while `none` yields to the opposing border. Set `DefaultTableBorders = true` only when you want a fallback grid on otherwise borderless tables.

Positioned tables in ordinary document flow preserve page, margin, or text anchors, explicit offsets, and text clearances. Following paragraphs use the available space beside the table and return to full width below it. Headings, lists, images, and other structured blocks move below an intersecting table. Positioned tables in multi-column sections retain an approximation warning.

### Export to bytes or streams

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("invoice.docx");

byte[] pdfBytes = document.ToPdfBytes();

using var stream = File.Create("invoice.pdf");
document.SaveAsPdf(stream);
```

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

- Paragraphs, headings, rich runs, links, bookmarks, page breaks, lists, and common spacing/indentation settings, including hanging and legal negative left/right indents.
- Word-authored text bullets use portable marker characters. Picture bullets currently use a text bullet in PDF output and report `NativePictureBulletTextFallback` with the source picture-bullet identifier; the embedded marker image is not rendered.
- Word sections, page size, orientation, margins, columns, headers, footers, page numbers, and document background color.
- Tables with common Word table styling, repeated headers, cell fills, borders, alignment, merged cells, and rich text in cells.
- Paragraph-aligned images, selected shapes, text boxes, content controls, simple form controls, footnote/endnote markers, and table-of-contents links where supported by the first-party PDF path.
- Unrotated, uncropped `InFrontOfText` images with explicit page-relative offsets inside the page bounds. Images follow the first page of their anchor paragraph or heading, including section columns, and paint over text and other flow content without reserving their height in the document flow. Overlapping foreground images follow `WordImage.ZOrder`; images with equal values retain document order.
- Per-operation conversion warnings through `PdfDocumentConversionResult.Report` or `PdfSaveResult.Report`.

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
| Convert | 1 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
