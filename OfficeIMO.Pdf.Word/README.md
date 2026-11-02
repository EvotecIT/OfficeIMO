# OfficeIMO.Pdf.Word

Reconstruct editable Word documents or visual page images from supported PDF content.

```sh
dotnet add package OfficeIMO.Pdf.Word
```

## Import semantic PDF content into Word

```csharp
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Pdf.Word;

PdfDocument pdf = PdfDocument.Load("packet.pdf");
using WordDocument word = pdf.ToWord(new WordPdfImportOptions(), out OfficeConversionReport report);
word.Save("packet.docx");

foreach (var warning in report.Diagnostics) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The semantic import path preserves document metadata, page breaks, headings, paragraphs, lists, detected run color/size/emphasis, logical tables, safe URI hyperlinks, supported internal destination links, embedded images, form-widget placeholders, and conversion diagnostics when those structures are available in the PDF logical model. `UseSharedPageReadingOrder` defaults to `true`, so semantic items follow the core engine's crop-, rotation-, spanning-band-, and column-aware order. Editable output also retains each source page's physical size, uses a 36-point page margin, suppresses extra paragraph spacing in dense source content, preserves detected table-column proportions, and positions supported axis-aligned images relative to the page. These defaults improve business-document reconstruction, but they do not turn semantic import into fixed-layout page recreation or Microsoft Word rendering parity.

Set `PreserveSourcePageSize`, `PreserveCompactSourceSpacing`, or `PreserveImagePlacementPosition` to `false` when normal Word flow is more useful than the corresponding source geometry. `EditablePageMarginPoints` controls the source-sized section margin. Source page sizes require `PreservePageBreaks = true` when more than one PDF page is imported.

For selected pages, Fast versus Structured reconstruction, or custom semantic budgets, pass the canonical read settings through the import options:

```csharp
using OfficeIMO;
OfficeConversionReport report = pdf.SaveAsWord(
    "packet-selection.docx",
    new WordPdfImportOptions {
        ReadOptions = new PdfLoadOptions {
            PageSelection = PdfPageSelection.Parse("1-3,5"),
            LayoutOptions = new PdfTextLayoutOptions { ForceSingleColumn = true },
            Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 5 }
        }
    });
```

`ReadOptions` is used only when the source is an opened `PdfDocument`. Accepted recognition content can be converted directly through the front door, preserving its detected tables and word geometry.

For table-only recovery, use the same façade with the explicit profile:

```csharp
using OfficeIMO;
pdf.SaveAsWord(
    "statement-tables.docx",
    WordPdfImportOptions.CreateTablesOnly());
```

## What it imports

- Parser-supported PDF metadata, page breaks, headings, paragraphs, lists, logical tables, safe URI hyperlinks, supported internal destination links, complete image-file payloads with transparency-mask fidelity metadata, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware soft-mask-capable simple `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware soft-mask-capable `Indexed` palette PDF image streams into editable `.docx` content when their filters are supported.
- Image fallback placeholders and form-widget placeholders with diagnostics instead of silently dropping unsupported objects.
- Page-range filtered imports through `PdfDocument.Read(new PdfLoadOptions { PageSelection = ... })`.
- Active hyperlink reconstruction for absolute `http`, `https`, and `mailto` URI annotations through `WordPdfImportOptions.ImportUriLinks` and `WordPdfImportOptions.AllowedHyperlinkUriSchemes`.
- Internal PDF destination reconstruction through `WordPdfImportOptions.ImportInternalLinks`, mapping supported page and named destinations to Word bookmarks and anchor hyperlinks.
- Native image embedding through `WordPdfImportOptions.ImportImages`; complete image files, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware soft-mask-capable simple 8-bit `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware soft-mask-capable `Indexed` palette streams are embedded when their filters are supported. Fully transparent and unplaced resources are suppressed. Images whose clips, placement soft masks, unresolved source transparency masks, or unsupported blend modes could expose hidden source pixels are not embedded as raw pictures; the report records a typed omission, and `WordPdfImportOptions.IncludeImagePlaceholders` can retain an editable marker instead. Supported opacity is mapped to native Word picture transparency.
- Per-operation import findings through `OfficeConversionReport.Diagnostics`. `HasLoss` and `ThrowIfLoss()` use typed approximation, omission, and failure evidence even when a diagnostic is informational. The table-only profile reports visible text outside detected tables as `pdf-text-content-not-imported` instead of treating an intentionally narrow extraction as lossless.

## Visual pages

For a PDF whose page appearance matters more than editability, use visual pages:

```csharp
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Word;

var pdf = PdfDocument.Load("source.pdf");
var options = WordPdfImportOptions.CreateVisualPages();
options.Dpi = 144;
options.ReadOptions = new PdfLoadOptions {
    PageSelection = PdfPageSelection.Parse("1-3,5")
};
pdf.SaveAsWord("visual-pages.docx", options);
```

Each selected page becomes an image on a Word section with the source page's physical dimensions. Text, links, and form fields are not editable in this mode. Rendering uses the managed PDF engine and reports its limitations. The default limits are 100 pages, 64 million pixels per page, and 256 MB of encoded page images; Word pages larger than 22 inches in either dimension are rejected. Use the default `EditableContent` mode to reconstruct supported text, tables, and images as Word objects.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Pdf.Word` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
