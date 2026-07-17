# OfficeIMO.Rtf.Pdf

First-party semantic RTF/PDF conversion for OfficeIMO.

This package converts between the semantic `OfficeIMO.Rtf` document model and the first-party `OfficeIMO.Pdf` document model. The RTF engine remains the lossless parse/edit/write layer; PDF export is a visual/content conversion to a fixed-layout format, while PDF import uses the first-party logical PDF reader to recover parser-supported metadata, pages, headings, grouped paragraphs, and list markers into an editable RTF document.

Supported export coverage includes semantic paragraphs, paragraph indentation/spacing/line-height/pagination controls, paragraph/style tab stops with supported alignment and leader mapping, section-owned blocks, section page breaks, page-starting section page setup, document and section page-border visual export, rich runs, list markers, document page setup, metadata, tables with horizontal and vertical merged-cell spans, repeating header rows, solid row/cell fills, cell padding, vertical alignment, side and diagonal cell borders, PNG/JPEG images, bookmarks, field result text, hidden text control, footnote/endnote/annotation bodies, and running header/footer text including first-page and even-page variants. RTF can model separate borders per page side; PDF export maps the first styled RTF page border to the first-party PDF engine's uniform page border decoration.

Supported import coverage includes PDF Info metadata, first-page paper size, logical headings, logical list items, grouped paragraphs, basic paragraph spacing, and page transitions as RTF page-break-before paragraphs. PDF is a fixed-layout format, so import is semantic text extraction rather than lossless visual reconstruction of arbitrary PDFs.

## Export with diagnostics

```csharp
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Pdf;

RtfDocument rtf = RtfDocument.Load("input.rtf").Document;
var options = new RtfPdfSaveOptions();
PdfDocumentConversionResult result = rtf.ToPdfDocumentResult(options);

result.Report.RequireNoErrorWarnings();
result.Save("output.pdf");
```

All PDF adapters expose the same `PdfResourcePolicy`. RTF conversion defaults to the portable deterministic policy and currently reads only content already present in the typed RTF model. A trusted-host policy is therefore explicit and future-proof rather than a hidden font-discovery default.

For raw RTF strings, bytes, or streams, use source-explicit APIs such as `ToPdfFromRtf()`, `ToPdfDocumentFromRtf()`, and `SaveAsPdfFromRtf()`. Typed `RtfDocument` instances use the standard `ToPdf()`, `ToPdfDocument()`, and destination-only `SaveAsPdf()` names.

PNG, JPEG, and supported DIB images use the shared managed drawing layer. Set `RtfPdfSaveOptions.ImageConverter` when WMF/EMF content must be rasterized; a null or invalid callback result is reported rather than silently treated as an image.

## Dependency footprint

- **External:** None by default. Applications may provide an optional WMF/EMF rasterization callback.
- **OfficeIMO:** `OfficeIMO.Rtf`, `OfficeIMO.Pdf`, and `OfficeIMO.Drawing` own semantic mapping, rendering, logical import, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
