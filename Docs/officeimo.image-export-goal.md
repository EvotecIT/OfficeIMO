# OfficeIMO Image Export Goal

This track turns OfficeIMO image export into a first-party capability across Excel, Word, PowerPoint, PDF, and shared drawing primitives.

## Non-Negotiable Dependency Rule

Product rendering paths stay dependency-free beyond the libraries OfficeIMO already owns for document semantics:

- OpenXML SDK for Office package structure.
- AngleSharp where HTML parsing/rendering adapters already use it.
- `OfficeIMO.Drawing` for shared pixels, SVG, PNG, paths, text layout, image projection, charts, diagnostics, and visual-quality helpers.

No product image export path may depend on external rendering products, Office automation, browser screenshots, native PDF rasterizers, Skia/ImageSharp/System.Drawing, or commercial document converters. External tools may exist only in tests, benchmarks, or manual comparison gates where they are explicitly not shipped product rendering.

## Current State

- Excel has the strongest PNG/SVG export path: ranges, worksheets, workbook batches, visual baselines, diagnostics, and tracked fidelity gaps.
- PowerPoint has a fixed-layout slide renderer with useful shape, picture, text, table, chart, and presentation batch coverage.
- Word has first-page PNG/SVG preview support and needs a real pagination model before multi-page export becomes production-grade.
- PDF owns PDF stream/page/writer behavior and uses Drawing for shared image and visual helpers, but first-party PDF page-to-image rendering still needs an internal content rendering plan.
- `OfficeIMO.Drawing` is the shared engine. Document packages should project source semantics into Drawing primitives, not grow private renderers.

## Execution Order

1. Keep the no-external-renderer contract enforced in tests.
2. Burn down Excel tracked visual-fidelity gaps while keeping Drawing reusable.
3. Add PowerPoint real-world fixture baselines because slides map cleanly to fixed Drawing scenes.
4. Build Word pagination in layers: paragraphs and tables, floating objects, then multi-page batch export.
5. Design PDF page-to-image as a first-party renderer over the PDF logical/content model, not a wrapper around PDFium, Poppler, browser screenshots, or Office automation.
6. Keep visual QA shared: approved PNG/SVG baselines, renderable/nonblank checks, and focused diagnostics for unsupported source features.

## Done Shape

The public surface should feel consistent:

```csharp
sheet.Range("A1:D12").SaveAsPng("range.png");
presentation.ToImages().ForSlideRange(1, 3).SaveAsSvg("slides");
document.ToImage().FirstPage().SaveAsPng("preview.png");
```

Future PDF page image APIs should follow the same builder/result/diagnostic model once the first-party renderer exists.
