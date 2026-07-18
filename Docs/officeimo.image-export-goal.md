# OfficeIMO Image Export Goal

This track keeps image export first-party and consistent across Drawing, Excel, Word, PowerPoint, HTML, OneNote, Visio, PDF, and source formats that already project into PDF.

## Non-Negotiable Dependency Rule

Product rendering paths stay dependency-free beyond the libraries OfficeIMO already owns for document semantics:

- OpenXML SDK for Office package structure.
- AngleSharp where HTML parsing/rendering adapters already use it.
- `OfficeIMO.Drawing` for shared pixels, PNG/JPEG/TIFF/SVG/WebP encoding, paths, text layout, image projection, charts, diagnostics, and visual-quality helpers.

No product image export path may depend on external rendering products, Office automation, browser screenshots, native PDF rasterizers, Skia/ImageSharp/System.Drawing, or commercial document converters. External tools may exist only in tests, benchmarks, or manual comparison gates where they are explicitly not shipped product rendering.

## Current State

- Excel has the strongest visual baseline path: ranges, worksheets, manual-page-break slices, workbook batches, diagnostics, and all five shared output formats.
- PowerPoint has a fixed-layout slide renderer with shape, picture, text, table, chart, presentation batch coverage, and representative fixture gates.
- Word has estimated multi-page rendering, page-range batch export, section-aware headers and footers, and all five shared output formats. It remains an OfficeIMO layout estimate rather than Microsoft Word's application-owned pagination.
- HTML renders continuous or paged surfaces through a shared Drawing scene, including synchronous and resource-aware asynchronous APIs.
- OneNote renders pages, sections, and notebooks with ink, math, image placeholders, batch selection, and bounded raster output.
- Visio keeps its established native SVG/raster geometry renderer and exposes format-neutral page and document batch export.
- PDF projects loaded pages into Drawing and exposes the same five-format single/batch contract, DPI/thumbnail controls, page selection, raster limits, and diagnostics as the Office document packages.
- `PdfDocumentConversionResult` is the single paged-image bridge for Markdown, AsciiDoc, LaTeX, RTF, OneNote, Word, Excel, PowerPoint, and HTML PDF adapters. Source conversion warnings flow into every image result.
- `OfficeIMO.Drawing` is the shared engine. Document packages project source semantics into Drawing primitives and reuse its result, options, safety, encoder, decoder, and diagnostic contracts.

## Execution Order

1. Keep result identity structural: `OfficeImageExportResult` rejects mismatched encoded formats or dimensions, and every converter uses it.
2. Keep raster limits pre-allocation: the Drawing planner combines caller, renderer, and encoder limits and applies one overflow policy.
3. Keep source decoding honest: bounded baseline TIFF and OfficeIMO's literal-lossless WebP subset are first-party; broader variants use a caller codec.
4. Expand representative visual baselines for PowerPoint, Word, HTML, OneNote, PDF, and adapters without copying renderer logic.
5. Extend PDF operator/font/form/transparency coverage in the existing first-party page-to-Drawing projection.
6. Add a source-specific direct image API only when it offers a visual contract that the shared `PdfDocumentConversionResult.ToImages()` bridge cannot represent.

## Done Shape

The public surface should feel consistent:

```csharp
sheet.Range("A1:D12").SaveAsPng("range.png");
presentation.ToImages().ForSlideRange(1, 3).AsSvg().Save("slides");
document.ToImage().FirstPage().AsPng().Save("preview.png");
html.ToImages().Paged().AsWebp().Save("pages");
visio.ToImages().AllPages().AsTiff().Save("diagram-pages");
PdfReadDocument.Load(pdfBytes).ToImages().Pages("1-3,last").AsWebp().Save("pdf-pages");
markdown.ToPdfDocumentResult().ToImages().AsPng().Save("markdown-pages");
```

All of these APIs return the same validated result shape, use portable batch filenames, and apply the same pre-allocation raster policy.
