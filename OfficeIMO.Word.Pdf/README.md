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

var options = new PdfSaveOptions {
    Orientation = PdfPageOrientation.Portrait,
    Margins = PageMargins.UniformCentimeters(1.5),
    Title = "Customer proposal",
    Author = "Evotec",
    IncludePageNumbers = true,
    PageNumberFormat = "Page {current} of {total}"
};

document.SaveAsPdf("proposal.pdf", options);
```

### Export to bytes or streams

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("invoice.docx");

byte[] pdfBytes = document.ToPdf();

using var stream = File.Create("invoice.pdf");
document.SaveAsPdf(stream);
```

### Capture conversion warnings without throwing away the report

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("complex-document.docx");
var options = new PdfSaveOptions {
    DefaultTableBorders = true
};

var result = document.TrySaveAsPdf("complex-document.pdf", options);
if (!result.Succeeded) {
    foreach (string diagnostic in result.Diagnostics) {
        Console.WriteLine(diagnostic);
    }
}

foreach (var warning in result.Warnings) {
    Console.WriteLine($"{warning.Source}: {warning.Message}");
}
```

### Import logical PDF tables into Word

```csharp
using OfficeIMO.Word.Pdf;

var imported = PdfWordTableConverterExtensions.SaveTablesAsWordDocument(
    "statement.pdf",
    "statement-tables.docx");

foreach (var table in imported) {
    Console.WriteLine($"Page {table.PageNumber}, table {table.TableIndex + 1}");
}
```

### Import semantic PDF content into Word

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

var options = new PdfWordReadOptions {
    LayoutOptions = new PdfTextLayoutOptions {
        ForceSingleColumn = true
    }
};

var import = File.ReadAllBytes("packet.pdf").ToWordDocumentFromPdfResult(options);
using WordDocument word = import.Value;
byte[] docx = word.ToBytes();

foreach (var warning in import.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The semantic import path preserves document metadata, page breaks, headings, paragraphs, lists, logical tables, safe URI hyperlinks, supported internal destination links, embedded images, form-widget placeholders, and conversion diagnostics when those structures are available in the PDF logical model. It creates an editable Word document; it does not claim fixed-layout page recreation or Microsoft Word rendering parity.

## What it exports

- Paragraphs, headings, rich runs, links, bookmarks, page breaks, lists, and common spacing/indentation settings.
- Word sections, page size, orientation, margins, columns, headers, footers, page numbers, and document background color.
- Tables with common Word table styling, repeated headers, cell fills, borders, alignment, merged cells, and rich text in cells.
- Paragraph-aligned images, selected shapes, text boxes, content controls, simple form controls, footnote/endnote markers, and table-of-contents links where supported by the first-party PDF path.
- Per-operation conversion warnings through `PdfDocumentConversionResult.Report` or `PdfSaveResult.Report`.

## What it imports

- Parser-supported PDF metadata, page breaks, headings, paragraphs, lists, logical tables, safe URI hyperlinks, supported internal destination links, complete image-file payloads with transparency-mask fidelity metadata, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware soft-mask-capable simple `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware soft-mask-capable `Indexed` palette PDF image streams into editable `.docx` content when their filters are supported.
- Image fallback placeholders and form-widget placeholders with diagnostics instead of silently dropping unsupported objects.
- Page-range filtered imports through `PdfWordReadOptions.PageRanges`.
- Active hyperlink reconstruction for absolute `http`, `https`, and `mailto` URI annotations through `PdfWordReadOptions.ImportUriLinks` and `PdfWordReadOptions.AllowedHyperlinkUriSchemes`.
- Internal PDF destination reconstruction through `PdfWordReadOptions.ImportInternalLinks`, mapping supported page and named destinations to Word bookmarks and anchor hyperlinks.
- Native image embedding through `PdfWordReadOptions.ImportImages`; complete image files, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware soft-mask-capable simple 8-bit `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware soft-mask-capable `Indexed` palette streams are embedded when their filters are supported. Pass-through JPEG image payloads with unresolved PDF transparency masks are embedded with `PdfImageTransparencyMaskNotResolved`; unsupported complex PDF image streams can still produce editable placeholders through `PdfWordReadOptions.IncludeImagePlaceholders`.
- Per-operation import warnings through `PdfWordConversionResult.Report`.

## Options and diagnostics

Use `PdfSaveOptions` when callers need to override page geometry, metadata, page-number behavior, font family, table-border fallback, profile presets, or text fallback policy. `TextFallbacks` uses the shared `PdfTextFallbackFeatures` enum. The balanced resource default enables installed fonts but denies arbitrary local and remote reads; use `PdfResourcePolicy.CreatePortableDeterministic()` for reproducible or untrusted conversion and `CreateTrustedHost()` only when local or remote resource access is intentional. Profiles do not inject page numbers; set `IncludePageNumbers = true` explicitly when generated numbering is desired. Request `ToPdfDocumentResult()` or `TrySaveAsPdf()` when diagnostics matter; unsupported Word features should become actionable operation results instead of mutable option state. Available embeddable Word families use shared named PDF resources and are not limited to three compatibility slots. Unavailable or non-embeddable families fall back to a mapped PDF font with an explicit warning.

## Boundaries

- This package does not try to be a full Word renderer with perfect Microsoft Word parity or a fixed-layout PDF-to-DOCX recreation engine.
- PDF-to-Word import is semantic reconstruction over parser-supported logical PDF objects. Complex/unsupported PDF image streams, interactive controls, unresolved destinations, and remote/cross-document PDF navigation actions are not yet reconstructed as native Word objects.
- Unsupported or simplified Word features should surface warnings rather than being hidden in the README as broad claims.
- Reusable PDF layout work belongs in `OfficeIMO.Pdf`; Word-specific mapping belongs here.
- PowerShell PDF workflows should be exposed through [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice).

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
- **OfficeIMO:** `OfficeIMO.Word`, `OfficeIMO.Pdf`, and `OfficeIMO.Drawing` own the source model, PDF engine, mapping, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
