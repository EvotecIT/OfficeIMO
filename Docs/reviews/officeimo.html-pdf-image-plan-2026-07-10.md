# OfficeIMO HTML to PDF and Image Implementation Plan

Date: 2026-07-10

Status: active implementation; initial end-to-end renderer slice complete

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

OfficeIMO has the shared foundation and an initial end-to-end vertical slice:

- `OfficeIMO.Html` parses HTML with AngleSharp, evaluates CSS media, resolves custom properties, discovers resources, enforces URL and document limits, and reports diagnostics.
- `OfficeIMO.Drawing` already supports the PNG/SVG rendering path used by Excel, Word, PowerPoint, and Visio.
- `OfficeIMO.Pdf` already supplies text, images, drawings, annotations, outlines, metadata, font embedding, extraction, security, and PDF inspection.
- `OfficeIMO.Html` now exposes continuous and paged render contracts plus direct PNG/SVG output through `OfficeIMO.Drawing`.
- `OfficeIMO.Html.Pdf` now has a `Rendered` profile that maps the shared render result directly to native PDF text, shapes, images, and link annotations. The existing semantic and document profiles remain available.

The current renderer deliberately starts with normal flow, styled text, grapheme-safe token fragmentation, non-BMP font cmap lookup, managed PDF font fallback and Latin-ligature controls, exact explicit-whitespace extraction, tables with row/column spans, images, generic `@page` geometry, named page assignment, first/left/right and named pseudo-page margin content, all standard page-margin box positions, repeated table header/footer groups, span-safe line/row pagination, and bounded asynchronous resource resolution. It is not yet a complete browser layout engine. Flex, grid, per-master page geometry, pseudo-page body reflow, advanced positioning, bidi, complex shaping, and other unfinished areas emit stable diagnostics or remain open below.

## Implemented Checkpoint

- [x] One shared render document and visual model in `OfficeIMO.Html` for continuous and paged output.
- [x] Direct PNG and SVG export through the existing dependency-free `OfficeIMO.Drawing` backend.
- [x] Direct searchable PDF output through a thin `OfficeIMO.Html.Pdf` adapter, without a PDF or image intermediate.
- [x] Synchronous APIs for embedded/local content and asynchronous APIs for application-resolved external resources.
- [x] URL policy, per-resource and total byte budgets, timeout, cancellation, content-type checks, surface limits, page limits, and layout-depth limits.
- [x] Screen/print media selection, CSS custom-property inheritance/fallback/cycle handling, basic page rules, and stable text/table fragmentation.
- [x] First/left/right page-margin content across all standard top, bottom, side, and corner boxes, with quoted text, page counters, font/color/alignment styling, and shared SVG/PDF output.
- [x] CSS `page` assignment plus named `@page` masters and named `:first`/`:left`/`:right` margin-content overrides, with the selected name retained on each rendered page.
- [x] Nested page fragmentation with widows/orphans, repeated table headers and footers, and parity-correct left/right/recto/verso breaks.
- [x] Table occupancy-grid layout for `colspan`, row-group-bounded `rowspan` including `rowspan="0"`, distributed span height, and span-safe page breaks.
- [x] Shared `OfficeIMO.Drawing` Unicode text-element measurement and HTML long-token fragmentation that preserve combining sequences and surrogate pairs, plus managed TrueType format-12 cmap lookup for non-BMP scalars.
- [x] Rendered PDF font fallback and shaping controls over the existing dependency-free `OfficeIMO.Pdf` engine, including caller-supplied embedded families/providers and exact Unicode whitespace extraction across rich text runs.
- [x] Public diagnostic-code catalog for implemented fallbacks and unsupported renderer behavior.
- [x] Contract tests for PNG, SVG, searchable PDF text, links, page geometry, media rules, custom properties, pagination, resources, timeout, cancellation, and diagnostics.

This checkpoint establishes the architecture and usable basic output. It does not close the remaining phases or justify full HTML/CSS fidelity claims.

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

- [x] Define the paged and continuous render-option contracts.
- [x] Define the initial shared rendered-document, page, shape, image, text-run, and link contracts.
- [ ] Add semantic-node, glyph, clipping, transform, and advanced paint contracts as their formatting phases land.
- [ ] Build a representative corpus: invoices, statements, reports, letters, certificates, catalog pages, dashboards, multilingual documents, and hostile-resource cases.
- [ ] Record expected geometry, page counts, extracted text, links, diagnostics, and visual baselines.
- [ ] Define performance and memory budgets by document class.

Exit gate: the public behavior and proof corpus exist before implementation details become permanent APIs.

### Phase 1 - Prepare the existing HTML foundation

- [x] Split `HtmlResourcePipeline.cs` into focused partials for element discovery, CSS discovery and syntax, custom-property resolution, selector matching, policy, and internal models.
- [x] Split `HtmlComputedStyleEngine.cs` by rules, media, supports, cascade, selector matching, CSS syntax, and internal models so the existing engine remains the shared owner.
- [x] Add cancellation and timeout propagation to direct-render resource loading and conversion orchestration.
- [ ] Carry the same cancellation contract through any shared discovery/loading paths reused outside the direct renderer.
- [x] Add the `OfficeIMO.Drawing` project reference to `OfficeIMO.Html`.
- [x] Establish initial shared units, coordinate systems, colors, font descriptors, and diagnostic codes.

Exit gate: parsing, CSS, resources, and diagnostics can serve Word/RTF conversion and the new renderer without duplicated logic.

### Phase 2 - Computed values and formatting tree

- [ ] Complete cascade, inheritance, media, and computed-value handling required by the corpus. Custom-property inheritance, fallback, and cycle handling are implemented.
- [ ] Build a typed formatting tree separate from the AngleSharp DOM.
- [ ] Resolve display roles, generated content, counters, replaced elements, pseudo-elements, and stacking contexts.
- [ ] Preserve source and semantic references needed for diagnostics, links, accessibility, and extraction order.

Exit gate: every rendered node has deterministic computed values or a stable unsupported-value diagnostic.

### Phase 3 - Typography and inline layout

- [ ] Consolidate font parsing, metrics, fallback, and glyph outlines in the existing shared owner, primarily `OfficeIMO.Drawing` where reusable. Grapheme-safe deterministic measurement, non-BMP TrueType cmap lookup, and rendered-PDF fallback wiring are implemented; shared HTML font discovery/fallback planning and richer OpenType metrics remain.
- [ ] Implement managed shaping, bidi resolution, ligatures, kerning, combining marks, and script-aware fallback. The direct PDF adapter now exposes the existing managed Latin-ligature mode and optional shaping-provider seam, but shared layout must still consume positioned glyph results.
- [ ] Implement whitespace, line breaking, inline boxes, decorations, baseline alignment, and intrinsic text measurement.
- [ ] Verify positioned glyphs round-trip to searchable/extractable PDF text. Current scalar text, Unicode fallback, and explicit spaces round-trip exactly; positioned shaped glyph sequences remain.

Exit gate: multilingual line boxes are deterministic across supported target frameworks and platforms.

### Phase 4 - Core layout

- [ ] Implement block flow, margin collapsing, padding, borders, backgrounds, sizing, and overflow.
- [ ] Implement replaced elements and image/SVG intrinsic sizing.
- [ ] Implement floats and relative/absolute/fixed/sticky positioning.
- [ ] Complete table layout and row/cell fragmentation. Occupancy-grid column placement, column spans, row-group-bounded row spans, span height, and span-safe row breaks are implemented; captions, intrinsic column sizing, border models, and cell-content fragmentation remain.
- [ ] Implement flexbox, grid, and multicolumn formatting contexts.

Exit gate: the continuous renderer can place the full target corpus without pagination.

### Phase 5 - Pagination and fragmentation

- [ ] Complete page construction and page masters. Generic `@page` size/margins, print media selection, named page assignment, and generic/named first/left/right margin-content selection are implemented; named and pseudo-page body-geometry reflow remains.
- [ ] Complete break rules and continuation state. Stable nested text/child/table-row continuation, CSS `widows`/`orphans`, parity-correct left/right/recto/verso breaks, repeated leading `<thead>` and trailing `<tfoot>` groups, and row-span-safe boundaries are implemented; advanced fragmentation remains.
- [ ] Complete generated page content, counters, running headers/footers, and margin boxes. Quoted text, `counter(page)`, `counter(pages)`, and all sixteen standard margin boxes are implemented; running strings, element content, and richer generated content remain.
- [ ] Make backgrounds, borders, positioned elements, flex/grid content, and links fragment correctly.

Exit gate: paged geometry and page counts match the accepted corpus.

### Phase 6 - Shared paint and semantic model

- [x] Convert the initial layout fragments into a backend-neutral ordered visual scene.
- [ ] Reuse or extend `OfficeIMO.Drawing` for paths, text outlines, images, gradients, transforms, clipping, shadows, and opacity.
- [ ] Keep searchable text runs and semantic nodes beside visual operations. Searchable text runs and links are implemented; semantic nodes remain.
- [ ] Complete deterministic z-order, clipping, resource reuse, and fallback diagnostics. Initial paint order and stable fallback diagnostics are implemented.

Exit gate: a single rendered result contains everything needed by both image and PDF backends.

### Phase 7 - HTML to PNG and SVG

- [x] Add `OfficeIMO.Html` image-export options aligned with `OfficeImageExportOptions`.
- [x] Add continuous `ToPng`, `ToSvg`, file, stream, synchronous, and asynchronous APIs.
- [x] Add paged `ExportImages` APIs with page numbering and diagnostics.
- [ ] Complete maximum surface, tiling, scale, DPI, transparency, and background behavior. Surface limits, scale, and background are implemented.
- [ ] Activate image baselines for both paged and continuous modes.

Exit gate: HTML image output uses only `OfficeIMO.Html` plus existing OfficeIMO projects and produces no PDF intermediate.

### Phase 8 - Direct HTML to PDF

- [x] Add a direct/paged `Rendered` profile in `OfficeIMO.Html.Pdf` that consumes the shared rendered-document model.
- [ ] Complete mapping to PDF structures. Searchable Unicode text with managed fallback controls, exact explicit spaces, basic shapes, images, and external links are implemented; positioned shaped glyphs and richer semantics remain.
- [x] Preserve existing semantic/document conversion profiles for users who want those contracts.
- [x] Add async/cancellable save APIs with explicitly buffered final PDF serialization.
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
| Buffered HTML-to-PDF orchestration | Async resource resolution and cancellation are implemented; keep final serialization explicitly documented as buffered until incremental PDF writing exists | Before claiming streaming output |
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
