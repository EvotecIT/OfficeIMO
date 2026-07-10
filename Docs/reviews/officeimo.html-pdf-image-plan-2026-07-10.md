# OfficeIMO HTML to PDF and Image Implementation Plan

Date: 2026-07-10

Status: proposed implementation plan

## Goal

Build one first-party HTML rendering pipeline that can produce:

- paged PDF with searchable text, links, outlines, metadata, and document semantics;
- paged PNG or SVG output, one image per rendered page;
- continuous PNG or SVG output for a configured screen viewport.

The renderer should ultimately cover the full business-document HTML/CSS surface, including modern layout and paged-media features. Delivery should be incremental, with every phase producing a useful and explicitly documented contract.

## Fixed Constraints

- Do not add new external dependencies.
- Continue using the existing `AngleSharp` and `AngleSharp.Css` dependencies in `OfficeIMO.Html`.
- Reuse `OfficeIMO.Drawing` for raster, SVG, drawing, font, and image primitives.
- Keep `OfficeIMO.Pdf` dependency-free and unaware of HTML, CSS, or AngleSharp.
- Do not create an `OfficeIMO.Html.PagedMedia` package. The HTML layout capability belongs in the existing `OfficeIMO.Html` package.
- Keep `OfficeIMO.Html.Pdf` as a thin PDF adapter over the shared HTML rendering model.
- Use one layout and paint model for PDF and image output. Do not render images by first generating a PDF, and do not create separate PDF and image layout engines.
- Unsupported or approximate behavior must produce stable diagnostics rather than silently changing the document.

## Ownership

| Project | Responsibility |
|---|---|
| `OfficeIMO.Html` | DOM parsing, resource loading and policy, CSS cascade, computed values, layout, pagination, diagnostics, and the shared rendered-document model |
| `OfficeIMO.Drawing` | Shared drawing scene, raster canvas, PNG encoder, SVG exporter, font measurement/outlines, clipping, gradients, transforms, and reusable paint primitives |
| `OfficeIMO.Html.Pdf` | Thin mapping from the shared rendered-document model to `OfficeIMO.Pdf` pages, text, drawings, images, links, annotations, outlines, metadata, and tags |
| `OfficeIMO.Pdf` | PDF composition, serialization, reading, editing, security, signatures, compliance, and document-native primitives |

`OfficeIMO.Html` may reference `OfficeIMO.Drawing`. That is reuse of an existing OfficeIMO-owned dependency-free project, not a new external dependency or a reason to introduce another package.

## Current State

OfficeIMO already has most of the surrounding pieces:

- `OfficeIMO.Html` parses HTML with AngleSharp, evaluates CSS media, computes a limited style summary, discovers resources, enforces URL and document limits, and reports diagnostics.
- `OfficeIMO.Drawing` already supports the PNG/SVG rendering path used by Excel, Word, PowerPoint, and Visio.
- `OfficeIMO.Pdf` already supplies text, images, drawings, annotations, outlines, metadata, font embedding, extraction, security, and PDF inspection.
- `OfficeIMO.Html.Pdf` already exposes HTML-to-PDF APIs and profile contracts, but the current paths translate through Markdown or Word instead of directly laying out authored HTML/CSS.

The missing shared capability is a direct HTML/CSS layout, pagination, and paint model. That is the component both HTML-to-image and authored-CSS HTML-to-PDF need.

## End-to-End Pipeline

```text
HTML + base URI + render options
              |
              v
AngleSharp DOM parsing and normalization
              |
              v
Resource discovery, policy, loading, and decoding
              |
              v
CSS cascade, inheritance, custom properties, and computed values
              |
              v
Formatting tree and intrinsic measurement
              |
              v
Block / inline / table / flex / grid / positioned layout
              |
              v
Continuous viewport or paged-media fragmentation
              |
              v
Shared visual scene + semantic/document information
          /                         \
         v                           v
OfficeIMO.Drawing                OfficeIMO.Html.Pdf
PNG / SVG                        searchable PDF
```

The shared result must retain more than pixels. It needs positioned text runs, font and glyph information, links, headings, alternative text, reading order, page geometry, clipping, and paint operations so each backend can preserve the strengths of its output format.

## Output Contracts

### Print and paged mode

- Uses print media rules.
- Honors page size, orientation, margins, page breaks, and `@page` rules.
- Produces one or more stable page results.
- Exports one PNG/SVG per page or one multi-page PDF.
- Supports repeated table headers and footers, running content, counters, widows, and orphans as those phases land.

### Screen and continuous mode

- Uses screen media rules.
- Accepts a viewport width and optional viewport height.
- Lays out the full document as one continuous surface.
- Produces a full-document PNG/SVG without inventing page breaks.
- Applies an explicit maximum-dimension and memory policy for very tall documents.

### Public API direction

The final names should follow existing OfficeIMO image-export conventions. The intended shape is:

```csharp
byte[] png = html.ToPng(options);
string svg = html.ToSvg(options);

IReadOnlyList<OfficeImageExportResult> pages =
    html.ExportImages(OfficeImageExportFormat.Png, pagedOptions);

PdfDocument pdf = html.ToPdfDocument(pdfOptions);
byte[] pdfBytes = html.SaveAsPdf(pdfOptions);
```

The image surface belongs to `OfficeIMO.Html`; the PDF extension surface remains in `OfficeIMO.Html.Pdf`. Async variants should be added for resource-loading and output workflows, with cancellation and timeout carried through the whole operation.

## Required Feature Contract

### HTML and resources

- HTML5 parsing, malformed-markup recovery, base URI handling, and nested documents.
- Inline, embedded, linked, and imported stylesheets.
- Images, data URIs, SVG, fonts, and controlled remote resources.
- Source-neutral URL policy, size/count/depth budgets, media-type validation, timeouts, and cancellation.
- Deterministic handling of unavailable or rejected resources.

### CSS values and cascade

- Origin, importance, specificity, source order, inheritance, and initial/unset/revert behavior.
- Custom properties and `var()` fallback/cycle handling.
- Absolute and relative lengths, percentages, viewport units, font-relative units, colors, and opacity.
- `calc()`, `min()`, `max()`, and `clamp()` where required by layout.
- `@media` for print and screen, `@supports`, counters, generated content, and font faces.

### Typography

- Font discovery, `@font-face`, fallback, embedding, and deterministic substitution diagnostics.
- Unicode line breaking, whitespace collapsing, tabs, soft hyphens, and optional managed hyphenation.
- Kerning, ligatures, combining marks, bidirectional text, complex scripts, vertical alignment, decorations, and letter/word spacing.
- Searchable PDF text and correct extraction order, not rasterized paragraphs.

Text shaping must be implemented with first-party managed code and existing OfficeIMO font primitives. The existing `IPdfTextShapingProvider` seam can remain useful, but this plan must not require a new HarfBuzz, Skia, or native-assets package.

### Layout

- Normal block and inline flow, anonymous boxes, margin collapsing, intrinsic sizing, and overflow.
- Replaced elements, aspect ratios, min/max sizing, and object fit/position.
- Floats, clearance, relative/absolute/fixed/sticky positioning, and stacking contexts.
- Tables with captions, spans, border models, intrinsic column sizing, and fragmentation.
- Flexbox, grid, and multicolumn layout.
- Transforms, clipping, backgrounds, borders, outlines, gradients, shadows, and opacity groups.

### Paged media

- Page size, orientation, margins, named pages, pseudo-pages, and page selectors.
- `break-before`, `break-after`, `break-inside`, legacy break aliases, widows, and orphans.
- Running headers and footers, margin boxes, page counters, generated content, and repeated table groups.
- Fragmentation of text, blocks, tables, flex/grid items, images, borders, and backgrounds.
- Fixed-position content repeated on applicable pages.

### PDF semantics

- Searchable text, font embedding/subsetting, Unicode maps, and logical extraction order.
- Internal and external links, destinations, bookmarks/outlines, metadata, and attachments where requested.
- Heading, paragraph, list, table, figure, and alternative-text semantics suitable for tagged-PDF work.
- Clear diagnostics when visual fidelity and semantic fidelity require different fallbacks.

## Implementation Order

### Phase 0 - Lock contracts and evidence

- [ ] Define the paged and continuous render-option contracts.
- [ ] Define the shared rendered-document, page, visual-operation, text-run, link, and semantic-node contracts.
- [ ] Build a representative corpus: invoices, statements, reports, letters, certificates, catalog pages, dashboards, multilingual documents, and hostile-resource cases.
- [ ] Record expected geometry, page counts, extracted text, links, diagnostics, and visual baselines.
- [ ] Define performance and memory budgets by document class.

Exit gate: the public behavior and proof corpus exist before implementation details become permanent APIs.

### Phase 1 - Prepare the existing HTML foundation

- [ ] Split `HtmlResourcePipeline.cs` by responsibility: discovery, CSS URL extraction, resolution, policy, loading, decoding, and reporting.
- [ ] Split/refactor `HtmlComputedStyleEngine.cs` so the existing cascade is extended rather than replaced by a second engine.
- [ ] Add cancellation and timeout propagation to resource discovery/loading and conversion orchestration.
- [ ] Add the `OfficeIMO.Drawing` project reference to `OfficeIMO.Html`.
- [ ] Establish shared units, coordinate systems, colors, font descriptors, and immutable diagnostics.

Exit gate: parsing, CSS, resources, and diagnostics can serve Word/RTF conversion and the new renderer without duplicated logic.

### Phase 2 - Computed values and formatting tree

- [ ] Complete cascade, inheritance, custom-property, media, and computed-value handling required by the corpus.
- [ ] Build a typed formatting tree separate from the AngleSharp DOM.
- [ ] Resolve display roles, generated content, counters, replaced elements, pseudo-elements, and stacking contexts.
- [ ] Preserve source and semantic references needed for diagnostics, links, accessibility, and extraction order.

Exit gate: every rendered node has deterministic computed values or a stable unsupported-value diagnostic.

### Phase 3 - Typography and inline layout

- [ ] Consolidate font parsing, metrics, fallback, and glyph outlines in the existing shared owner, primarily `OfficeIMO.Drawing` where reusable.
- [ ] Implement managed shaping, bidi resolution, ligatures, kerning, combining marks, and script-aware fallback.
- [ ] Implement whitespace, line breaking, inline boxes, decorations, baseline alignment, and intrinsic text measurement.
- [ ] Verify positioned glyphs round-trip to searchable/extractable PDF text.

Exit gate: multilingual line boxes are deterministic across supported target frameworks and platforms.

### Phase 4 - Core layout

- [ ] Implement block flow, margin collapsing, padding, borders, backgrounds, sizing, and overflow.
- [ ] Implement replaced elements and image/SVG intrinsic sizing.
- [ ] Implement floats and relative/absolute/fixed/sticky positioning.
- [ ] Implement table layout and row/cell fragmentation prerequisites.
- [ ] Implement flexbox, grid, and multicolumn formatting contexts.

Exit gate: the continuous renderer can place the full target corpus without pagination.

### Phase 5 - Pagination and fragmentation

- [ ] Add page construction, page masters, `@page`, named pages, and print media selection.
- [ ] Add break rules, widows/orphans, continuation state, and repeated table groups.
- [ ] Add generated page content, counters, running headers/footers, and margin boxes.
- [ ] Make backgrounds, borders, positioned elements, flex/grid content, and links fragment correctly.

Exit gate: paged geometry and page counts match the accepted corpus.

### Phase 6 - Shared paint and semantic model

- [ ] Convert layout fragments into a backend-neutral ordered visual scene.
- [ ] Reuse or extend `OfficeIMO.Drawing` for paths, text outlines, images, gradients, transforms, clipping, shadows, and opacity.
- [ ] Keep searchable text runs and semantic nodes beside visual operations.
- [ ] Add deterministic z-order, clipping, resource reuse, and fallback diagnostics.

Exit gate: a single rendered result contains everything needed by both image and PDF backends.

### Phase 7 - HTML to PNG and SVG

- [ ] Add `OfficeIMO.Html` image-export options aligned with `OfficeImageExportOptions`.
- [ ] Add continuous `ToPng`, `ToSvg`, save, stream, and builder APIs.
- [ ] Add paged `ExportImages` APIs with page numbering and per-page diagnostics.
- [ ] Add maximum surface, tiling, scale, DPI, transparency, and background behavior.
- [ ] Activate image baselines for both paged and continuous modes.

Exit gate: HTML image output uses only `OfficeIMO.Html` plus existing OfficeIMO projects and produces no PDF intermediate.

### Phase 8 - Direct HTML to PDF

- [ ] Change the direct/paged profile in `OfficeIMO.Html.Pdf` to consume the shared rendered-document model.
- [ ] Map text runs to PDF text, visual operations to PDF drawing primitives, and links/semantics to PDF structures.
- [ ] Preserve existing semantic/document conversion profiles for users who want those contracts.
- [ ] Add async/cancellable save APIs with honest buffering behavior.
- [ ] Validate page geometry, extraction, links, outlines, metadata, encryption, and tagged structure.

Exit gate: PDF and image output agree on layout while PDF preserves text and document semantics.

### Phase 9 - Advanced fidelity

- [ ] Complete advanced SVG, filters, masks, blend modes, complex shadows, and raster fallbacks.
- [ ] Complete difficult table/flex/grid/multicolumn fragmentation cases.
- [ ] Expand modern CSS value/functions and generated-content behavior.
- [ ] Add managed hyphenation dictionaries only if they can be shipped without a new runtime dependency and with acceptable package size.

Exit gate: every accepted feature has corpus proof and every unimplemented feature has a stable diagnostic.

### Phase 10 - Hardening and release readiness

- [ ] Add fuzz/hostile-input, resource-budget, timeout, cancellation, and deterministic-output tests.
- [ ] Add executable NativeAOT and trimming smoke tests for the dependency-free PDF core and compatible renderer surfaces.
- [ ] Benchmark parse, style, layout, paint, image export, and PDF export independently.
- [ ] Publish a generated support matrix from profile contracts and diagnostic catalogs.
- [ ] Document platform/target-framework differences and memory limits.

Exit gate: support claims are generated from passing evidence, not maintained as an aspirational list.

## Existing Issues to Close

These are part of the roadmap, but they should not distort the dependency order of the renderer.

| Issue | Required action | When it blocks |
|---|---|---|
| Legacy PDF encryption output | Add a modern standard-security writer, prefer AES-256, make modern encryption the default, and keep RC4 only as an explicit legacy option if required | Before calling direct HTML-to-PDF production-ready for sensitive documents; can be implemented in parallel with layout |
| No first-party complex text shaping | Build managed shaping and shared font behavior from existing OfficeIMO primitives; do not add native or external packages | Before Phase 3 exits and before multilingual fidelity claims |
| Synchronous HTML-to-PDF orchestration | Add cancellation-aware async resource and conversion APIs; state clearly when final output is buffered | Before public renderer APIs stabilize |
| Oversized HTML resource/style files | Split by semantic responsibility and retain one shared CSS/resource engine | Before substantial renderer code is added |
| Incomplete computed-style surface | Extend the existing style engine with typed computed values and diagnostics | Before layout phases can be correct |
| No active end-to-end visual proof | Commit real paged and continuous baselines plus geometry/text assertions | Before fidelity claims or release |
| AOT/trimming proof is analyzer-heavy | Add executable NativeAOT smoke coverage where the actual dependency graph permits it | Before publishing AOT guarantees |
| Support information is scattered | Generate the public matrix from profile contracts, diagnostics, and passing corpus cases | Before declaring the feature generally available |

## Guardrails

- Do not add a browser process, headless browser, JavaScript runtime, native graphics stack, or new HTML/CSS parser.
- Do not add a new package merely to hold layout code; use focused folders, types, and namespaces inside `OfficeIMO.Html`.
- Do not put HTML/CSS concerns in `OfficeIMO.Pdf` or duplicate PDF serialization inside `OfficeIMO.Html`.
- Do not build separate layout code for PNG, SVG, and PDF.
- Do not rasterize whole pages for PDF except as an explicit, diagnosed fallback for a feature that cannot yet be expressed natively.
- Do not promise streaming until PDF objects are emitted incrementally; async writing of a completed byte array is still buffered output.
- Do not publish result fields, timings, or support claims until they are populated and tested.
- Split formatting contexts and pagination responsibilities into focused types; avoid another monolithic layout engine.

## Definition of Done

The end-to-end goal is complete when:

- the same HTML source and options produce matching PNG/SVG and PDF geometry from one layout result;
- both print/paged and screen/continuous modes have stable public contracts;
- PDF output remains searchable, link-aware, semantically structured, and compatible with existing OfficeIMO.Pdf features;
- PNG/SVG output follows established OfficeIMO.Drawing export patterns;
- complex scripts, tables, flexbox, grid, positioning, paged media, and advanced visuals are either proven or explicitly diagnosed;
- untrusted resources are bounded, cancellable, and policy-controlled;
- no new external dependency or package is required;
- the compatibility matrix is generated from passing corpus, visual, geometry, extraction, and interoperability tests.

The architectural center is the reusable renderer in `OfficeIMO.Html`: parse once, compute once, lay out once, and project the result to image or PDF without creating another package or another rendering brain.
