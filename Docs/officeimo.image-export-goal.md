# OfficeIMO Image Export Goal

This track keeps image export first-party and consistent across Drawing, Excel, Word, PowerPoint, HTML, OneNote, Visio, and future PDF page rendering.

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
- PDF owns PDF stream/page/writer behavior and uses Drawing for shared image and visual helpers, but first-party PDF page-to-image rendering still needs an internal content rendering plan.
- `OfficeIMO.Drawing` is the shared engine. Document packages should project source semantics into Drawing primitives, not grow private renderers.

## Execution Order

1. Enforce a format-conformance contract: every result's declared format must match its encoded bytes.
2. Standardize raster pixel budgets and encoder-dimension handling across document packages.
3. Add TIFF and WebP source decoding in Drawing when test fixtures justify the implementation.
4. Expand representative visual baselines for PowerPoint, Word, HTML, and OneNote without copying renderer logic.
5. Design PDF page-to-image as a first-party renderer over the PDF logical/content model, not a wrapper around PDFium, Poppler, browser screenshots, or Office automation.
6. Add thin direct adapters for formats such as Markdown and RTF only where an existing native model or HTML projection preserves useful diagnostics.

## Done Shape

The public surface should feel consistent:

```csharp
sheet.Range("A1:D12").SaveAsPng("range.png");
presentation.ToImages().ForSlideRange(1, 3).AsSvg().Save("slides");
document.ToImage().FirstPage().AsPng().Save("preview.png");
html.ToImages().Paged().AsWebp().Save("pages");
visio.ToImages().AllPages().AsTiff().Save("diagram-pages");
```

Future PDF page image APIs should follow the same builder/result/diagnostic model once the first-party renderer exists.
