# OfficeIMO HTML to PDF and Image Implementation Plan

Date: 2026-07-11

Status: end-to-end architecture and prioritized implementation checkpoint complete; advanced fidelity roadmap active

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
- Keep paged-media layout in the existing `OfficeIMO.Html` package; it is one rendering mode of the shared HTML engine, not a separate package boundary.
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

OfficeIMO has the shared foundation and a broad end-to-end implementation checkpoint:

- `OfficeIMO.Html` parses HTML with AngleSharp, evaluates CSS media, resolves custom properties, discovers resources, applies policy-approved external stylesheets and recursive imports in document order, enforces URL and document limits, and reports diagnostics.
- `OfficeIMO.Drawing` supplies the PNG/SVG rendering path used by Excel, Word, PowerPoint, and Visio, plus scoped caller-supplied TrueType faces shared by measurement, rasterization, and portable SVG embedding.
- `OfficeIMO.Pdf` already supplies text, images, drawings, annotations, outlines, metadata, font embedding, extraction, security, and PDF inspection.
- `OfficeIMO.Html` now exposes continuous and paged render contracts plus direct PNG/SVG output through `OfficeIMO.Drawing`.
- `OfficeIMO.Html.Pdf` now maps the shared render result directly to native PDF text, shapes, images, and link annotations. HTML-to-PDF has one direct path; Word and Markdown projections are explicit target conversions in their owning packages.

The current renderer deliberately starts with normal flow, block-level horizontal and vertical flex layout with wrapping, atomic inline flex boxes, block and atomic inline grid with explicit/implicit and responsive auto-fit/auto-fill tracks, row/column auto-flow, and named areas/lines, bounded multicolumn layout with balanced or explicit-height fill, full-width direct block spanners, vector column rules, and safe shared page-break boundaries, left/right and direction-aware logical-side floats with line-band wrapping and clearance, per-axis `visible`/`hidden`/`clip`/`auto`/`scroll` overflow groups with document-root/body viewport propagation, paint-only relative positioning, out-of-flow absolute positioning against initial, block, flex, grid, and measured inline containing rectangles, block/flex/grid/inline static-position anchoring for automatic insets, grid-area containing rectangles, fixed overlays repeated at viewport coordinates, negative/auto/positive numeric stacking bands across absolute, fixed, relative, and sticky block/flex/grid/inline contexts, block and atomic inline 2D transform/opacity stacking contexts, stable static-document snapshots for sticky content, styled source text and `::before`/`::after` text with attributes and scoped counters, grapheme-safe token fragmentation, simple RTL/Hebrew positioning with logical shared-scene order and PDF `ActualText`, dependency-free contextual forms for the core Arabic alphabet, non-BMP font cmap lookup, policy-bound TrueType `@font-face` activation, managed PDF font fallback and Latin-ligature controls, exact explicit-whitespace extraction, policy-bound external stylesheet/import loading, preserved CSS font-family fallback lists, configurable source-character and DOM-node budgets before style/layout work, tables with row/column spans, intrinsic/aspect-sized block/flex/grid/float/normal-inline images with all five object-fit modes, positioned source crops, wrapping, baseline participation, and clipped authored-size continuation when an image or non-rectangular paint path is taller than a page, layered CSS URL backgrounds with clipped `cover` and all standard repeat modes, opaque multi-stop linear gradients with hard edges and length/percentage stops, centered and off-center opaque multi-stop radial circles and ellipses with keyword, length, or percentage geometry/stops, and root-canvas propagation, generic `@page` geometry, named page assignment, first/left/right and named pseudo-page margin content, all standard page-margin box positions, repeated table header/footer groups, span-safe line/row pagination, and bounded asynchronous resource resolution. It is not yet a complete browser layout engine. Interactive scrollbars, cross-block float intrusion, `shape-outside`, full inline flex/grid baseline synthesis, normal-inline transform/opacity grouping, 3D and perspective transforms, blend modes, inline background/border fragment synthesis, nested generated flex/grid formatting contexts, subgrid, masonry, nested descendant column spanners and advanced column-rule styles, generated images and quote-state tokens, conic, repeating, and alpha-varying gradients, other CSS URL paint assets, WOFF/WOFF2 and CFF decoding, per-master page geometry, pseudo-page body reflow, explicit bidi embedding/isolate controls, joining scripts outside the bounded Arabic shaper, full OpenType shaping, and other unfinished areas emit stable diagnostics or remain open below.

## Implemented Checkpoint

- [x] One shared render document and visual model in `OfficeIMO.Html` for continuous and paged output.
- [x] Direct PNG and SVG export through the existing dependency-free `OfficeIMO.Drawing` backend.
- [x] Direct searchable PDF output through a thin `OfficeIMO.Html.Pdf` adapter, without a PDF or image intermediate.
- [x] Synchronous APIs for embedded/local content and asynchronous APIs for application-resolved external resources.
- [x] Policy-, timeout-, cancellation-, media-type-, byte-, count-, and import-depth-controlled external stylesheet graphs in DOM cascade order, including recursive imports, cycle suppression, media conditions, external `@page` rules, and explicit diagnostics for synchronous or unsupported URL-paint gaps.
- [x] URL policy, per-resource and total byte budgets, timeout, cancellation, content-type checks, surface limits, page limits, layout-depth limits, per-element background-layer limits, per-gradient color-stop limits, and operation-wide background-tile limits.
- [x] Screen/print media selection, CSS custom-property inheritance/fallback/cycle handling, basic page rules, and stable text/table fragmentation.
- [x] First/left/right page-margin content across all standard top, bottom, side, and corner boxes, with quoted text, page counters, font/color/alignment styling, and shared SVG/PDF output.
- [x] CSS `page` assignment plus named `@page` masters and named `:first`/`:left`/`:right` margin-content overrides, with the selected name retained on each rendered page.
- [x] Nested page fragmentation with widows/orphans, repeated table headers and footers, and parity-correct left/right/recto/verso breaks.
- [x] Table occupancy-grid layout for `colspan`, row-group-bounded `rowspan` including `rowspan="0"`, distributed span height, and span-safe page breaks.
- [x] Shared `OfficeIMO.Drawing` Unicode text-element measurement and HTML long-token fragmentation that preserve combining sequences and surrogate pairs, plus managed TrueType format-12 cmap lookup for non-BMP scalars.
- [x] Scoped TrueType `@font-face` loading from inline, data-URI, linked, and recursively imported CSS under the shared URL/count/depth/byte policy; the same validated faces drive HTML measurement, PNG rasterization, embedded SVG fonts, and rendered-PDF embedding. Unsupported WOFF/WOFF2/CFF inputs are diagnosed without adding codecs.
- [x] Layered CSS background-image loading from inline, data-URI, linked, and recursively imported CSS, including per-layer size, position, clipped `cover`, `repeat`, `space`, `round`, `repeat-x`, `repeat-y`, root-canvas propagation, normal boxes, and table cells. Layers paint back-to-front, `none` retains CSS list semantics without blocking body-canvas propagation, and paint is bounded per element and operation. Raster sources use one shared pattern visual while the thin PDF adapter expands clipped placements over deduplicated image resources. Supported SVG sources parse once into a shared Drawing scene and expand bounded vector tiles under one rectangular/rounded clip, sharing the layout snapshot across tiles and keeping PNG, SVG, and PDF native. Bounded opaque multi-stop `linear-gradient()` and centered/off-center circular or elliptical `radial-gradient()` layers with keyword geometry, percentage/absolute/font-relative stop positions, one/two-position color stops, duplicate hard stops, and backward-position fix-up share native Drawing gradients and PDF stitching functions; PDF serializes hard edges through a microscopic strictly ordered stitching interval. Other gradient forms still diagnose.
- [x] Backend-neutral clipped visual groups for per-axis `overflow`, `overflow-x`, and `overflow-y` values `visible`, `hidden`, `clip`, `auto`, and `scroll`. Generic, positioned, flex, grid, nested, and paginated content keeps vector shapes, native images, searchable text, clipped PDF links, and identical PNG/SVG/PDF geometry. `overflow-clip-margin` supports content/padding/border visual-box origins plus a non-negative absolute outset only for `clip`, as specified; scrollable boxes use a diagnosed initial static snapshot without fake interactive controls.
- [x] CSS `position: relative` for block, image, table, rule, flex/grid items, and nested inline content. Leading insets win over trailing insets, horizontal percentages resolve against the containing width, vertical percentages resolve for explicitly sized containing blocks, and paint offsets do not change line wrapping, flow height, sibling placement, or page-fragment ownership. Block, root, flex, grid, and inline relative/sticky contexts share numeric negative/auto/positive stacking with absolute/fixed layers while nested contexts remain atomic.
- [x] Out-of-flow `position: absolute` for the initial containing block, nested positioned block/inline ancestors, and direct flex/grid children. Explicit and opposing insets resolve against the active containing rectangle, percentages remain deterministic, and automatic insets use hypothetical flow markers accumulated through nested blocks, flex items, grid items, and inline fragments. Direct flex children honor main/cross alignment; direct grid children use declared numeric/named grid areas; positioned inline ancestors use measured wrapped-fragment bounds. Nested descendants are extracted without consuming flow or line width, and the same positioned visuals feed PNG, SVG, and searchable PDF output. Negative layers paint behind in-flow content but above the containing background; auto/zero and positive integer layers paint in numeric then source order across out-of-flow and in-flow positioned contexts; nested descendants cannot escape their parent stacking context.
- [x] Left/right and direction-aware `inline-start`/`inline-end` floats inside inline formatting content, including shrink-to-fit text boxes, intrinsic-ratio images, same-side packing, side-specific `clear`, variable line bands, nested inline discovery, full-width restoration below floats, and shared PNG/SVG/searchable-PDF output. Unsupported values use a cataloged diagnostic; cross-block intrusion, `shape-outside`, and advanced float fragmentation remain open.
- [x] `position: fixed` content is removed from flow and repeated at viewport coordinates on every paged surface or once on a continuous surface. `position: sticky` remains in flow as a stable static-document snapshot and emits an informational diagnostic because scroll state has no meaning in exported output.
- [x] Styled `::before` and `::after` text from quoted strings, `attr()`, `counter()`, and `counters()`, including scoped `counter-reset`, `counter-set`, and `counter-increment`, decimal/leading-zero/alphabetic/Roman styles, legacy single-colon syntax, CSS escapes, custom properties, link ownership, block-container placement, and stable diagnostics for unsupported content or counter declarations.
- [x] Block-level `row` and `row-reverse` flex layout with `nowrap`, `wrap`, and `wrap-reverse`; ordered element items; per-line `flex` grow/shrink/basis allocation; min/max freezing and redistribution; row/column gaps; main-axis distribution; cross-axis item and line alignment/stretch; nested flex containers; relative-positioned items; stable order-modified paint; and explicit fallback diagnostics for unsupported flex cases and values.
- [x] Block-level `column` and `column-reverse` flex layout with `nowrap`, `wrap`, and `wrap-reverse`, including definite-height percentage bases, auto-height percentage fallback, intrinsic cross-size planning, per-column vertical grow/shrink/basis and min/max constraints, row/column gaps, main and cross distribution, horizontal alignment/stretch with reversed flex cross-start, order-modified reverse placement, complete-item page breaks for a single column, and shared PNG/SVG/searchable-PDF output.
- [x] Flex item construction for direct anonymous text, styled generated `::before`/`::after` text, recursively flattened `display: contents`, blockified inline children, inherited link ownership, and four-sided auto margins that absorb main/cross free space before alignment. `inline-flex` now participates in inline layout as a shrink-to-fit atomic box with shared child visuals and a box-level link area; full typographic baseline synthesis remains.
- [x] Block and atomic inline grid with bounded integer, `auto-fill`, and `auto-fit` `repeat()` expansion; fixed, percentage, `auto`, `fr`, and `minmax()` tracks; explicit and implicit rows/columns; trailing empty-track collapse for auto-fit; sparse/dense row and column auto-flow; numeric and named lines; rectangular named template areas; spans; per-item/container alignment and stretch; anonymous/generated/`display: contents` items; auto margins; inherited atomic link areas; span-aware row break boundaries; stable value and limit diagnostics; and shared PNG/SVG/searchable-PDF output.
- [x] Bounded multicolumn formatting with `column-count`, `column-width`, the `columns` shorthand, normal and explicit gaps, balanced and explicit-height `column-fill`, direct block `column-span: all` partitions, solid/dashed/dotted/double vector column rules, overflow-column limits, shared safe fragmentation boundaries, generated-content ordering, and PNG/SVG/searchable-PDF parity.
- [x] Styled table captions participate in the table wrapper above or below the grid through inherited `caption-side: top | bottom`. Caption margins, padding, border, background, inline wrapping, empty-table retention, semantic text, page ownership, and shared PNG/SVG/searchable-PDF output use the same scene; invalid values diagnose a stable top fallback.
- [x] Table columns use bounded intrinsic text and replaced-image measurement for `table-layout:auto`, including min-content tokens, preferred unwrapped content, image natural/authored constraints, cell width constraints, colspans, and `<col>`/`<colgroup>` declarations. `table-layout:fixed` honors declared column and first-row cell widths, then distributes remaining width deterministically; cell placement, spans, wrapping, page breaks, and all output backends consume the same resolved track geometry.
- [x] Separate-border tables honor inherited one/two-axis non-negative absolute `border-spacing` around and between tracks, including rowspans and page-safe row offsets. `border-collapse:collapse` suppresses spacing and resolves each shared edge once across table, column-group, column, row-group, row, and cell origins. Hidden and none handling, width, supported style, origin, and deterministic source-order precedence select one shared Drawing line for PNG, SVG, and PDF; collapsed table borders are not painted a second time through the normal box path.
- [x] Backend-neutral isolated paint groups for block, shrink-to-fit `inline-block`, `inline-flex`, and `inline-grid` 2D CSS `matrix`, translate, scale, rotate, and skew transforms with two-dimensional transform origins and percentage opacity. The same nested scene drives dependency-free raster composition, SVG affine groups, and searchable PDF transparency Form XObjects; links and paged fragments retain transformed geometry. Unsupported normal-inline and non-2D effects emit stable diagnostics.
- [x] Full one-to-four-corner circular/elliptical CSS `border-radius` geometry with slash-separated axes, corner longhands, and CSS overlap normalization. Uniform cases retain compact native rounded rectangles; asymmetric/elliptical color, gradient, border, outline, shadow-carrier, URL-image, repeated-pattern, and replaced-image paint use one shared cubic path across PNG, SVG, and PDF. Independent border-side widths, colors, solid/dashed/dotted/double styles, layout insets, and half-corner paths use the same backend-neutral geometry.
- [x] Normal-flow vertical block margins collapse using the CSS positive/negative margin-set rule across adjacent siblings, parent/first-child and parent/last-child boundaries, and zero-height empty blocks. Padding, borders, explicit/minimum height, formatting-context boundaries, and inline content stop collapse. External collapsed margins propagate through nested flow metadata so paint placement, positioned static origins, page-break offsets, PNG/SVG coordinates, and paged PDF fit decisions agree. Browser-specific root/body margin quirks remain outside this normal-flow contract.
- [x] Viewport overflow resolves from the document root when it declares non-visible overflow and otherwise falls back to the body. One backend-neutral per-axis clip wraps normal, positioned, and fixed content after stacking while canvas backgrounds remain outside; continuous SVG/PNG and every PDF page use the same propagated viewport geometry.
- [x] Layered outer and inset CSS `box-shadow` with offsets, optional bounded approximate blur, signed spread, current/supported color, alpha, CSS front-to-back paint order, resolved per-corner geometry, and a configurable per-element layer limit. Outer shadows reuse shared Drawing shadow carriers; inset shadows use shared even-odd Drawing paths and rounded clips, so PNG, SVG, and PDF consume the same visual model.
- [x] Replaced block/flex/grid/float/normal-inline image sizing from intrinsic pixel dimensions and DPI, authored width/height, `box-sizing`, bounded positive `aspect-ratio`, and min/max constraints. `object-fit: fill | contain | cover | none | scale-down` plus keyword, length, percentage, and edge-offset `object-position` values resolve to one backend-neutral placement and source crop shared by PNG, SVG, and PDF. Per-corner content clipping, normal-inline wrapping, inherited links, and deterministic baseline participation reuse the same image box.
- [x] Shared SVG image metadata resolves absolute CSS lengths, pica and quarter-millimeter units, positive `viewBox` geometry, and intrinsic aspect ratios. A single authored width or height derives its missing partner from `viewBox` without discarding the authored dimension; relative dimensions fall back to the intrinsic `viewBox` size. A bounded dependency-free `OfficeIMO.Drawing` scene reader activates rect, rounded-rect, circle, ellipse, line, polygon, polyline, path, and text primitives in `<img>` and CSS background sources. Paint supports solid colors plus local `url(#id)` linear/radial fill and stroke servers in object-bounding-box or user-space coordinates, including percentages, bounded stop colors/opacities, definition-tree `currentColor`, shape-relative deferred resolution, bounded cycle-safe same-type `href` inheritance, arbitrary affine linear `gradientTransform`, exact translated/axis-scaled radial transforms, and bounded linear `repeat`/`reflect` expansion into one shared multi-stop gradient. Local `<use>` references expand ordinary shapes and groups with inherited paint, x/y placement, transforms, duplicate-ID rejection, cycle/depth protection, and the shared element-operation budget. Local `<symbol viewBox>` references additionally map explicit/default width and height through every standard `preserveAspectRatio` alignment, `meet`, `slice`, or `none` mode and clip content to the resulting viewport. Path data covers absolute/relative move, line, horizontal, vertical, cubic, smooth cubic, quadratic, smooth quadratic, elliptical arc, and close commands under a command limit. Endpoint arcs normalize radii and convert rotated large/small clockwise/counter-clockwise geometry to shared cubic paths with an exact final endpoint. Ordered nested SVG transform attributes support matrix, translate, scale, rotate around the origin or a center, and x/y skew in normalized `viewBox` coordinates. Text supports bounded nested `<tspan>` runs, inherited family/size/weight/style/solid fill/opacity, first-value x/y/dx/dy positioning, whitespace normalization or `xml:space`, start/middle/end anchored chunks, bounded `textLength` with `spacingAndGlyphs`, and composed arbitrary affine transforms through shared Drawing effect groups. PNG and SVG consume Drawing text directly; the thin PDF adapter recursively projects affine Drawing groups to native searchable text in scene order and discovers active embedded font families through nested rectangular, path-clip, and effect groups. The shared drawing visual scales and clips the whole scene through all output paths; CSS backgrounds reuse it for all supported sizing, positioning, and repeat modes under the existing tile budget. DTD/entity input, external or ambiguous references, reference cycles, oversized stop/run lists, rotated/sheared or repeating radial paint servers, and omitted features diagnose. Gradient text, per-glyph coordinate lists, spacing-only length adjustment, text paths, masks, filters, and CSS-style transforms remain open.
- [x] Backend-neutral path-clipped visual groups over existing `OfficeIMO.Drawing.OfficeClipPath` geometry, including direct rounded/freeform canvas clipping in `OfficeIMO.Pdf`. Replaced images and single/repeated CSS URL backgrounds now retain native encoded images inside the same rounded clip for PNG, SVG, and PDF instead of rasterizing or using rectangular fallback.
- [x] Rendered PDF font fallback and shaping controls over the existing dependency-free `OfficeIMO.Pdf` engine, including caller-supplied embedded families/providers and exact Unicode whitespace extraction across rich text runs.
- [x] Public diagnostic-code catalog for implemented fallbacks and unsupported renderer behavior.
- [x] Contract tests for PNG, SVG, searchable PDF text, links, page geometry, media rules, custom properties, generated content and counters, floats and clearance, sibling/parent/empty positive-and-negative block-margin collapse and collapse boundaries, per-axis overflow clipping, visual-box clip margins, and static scroll snapshots, relative/absolute/fixed/sticky positioning, block/flex/grid/wrapped-inline containing rectangles and static anchors, positioned grid areas, nested numeric and transform/opacity stacking contexts, isolated descendant and gradient opacity, affine scale/rotation, uniform/asymmetric/elliptical rounded color/border/gradient/URL/replaced-image boxes, uniform and side-specific solid/dashed/dotted/double borders, offset outlines, layered outer/inset alpha/blur/spread shadows, intrinsic/aspect/min-max image sizing, all five object-fit modes, object-position placement/cropping, oversized raster/SVG page continuation, table captions, spans, intrinsic columns, border spacing, and collapsed-border conflict origins, fixed page repetition, horizontal/vertical/inline flex, block/inline grid, row/column and dense grid flow, named areas/lines, wrapping, multicolumn balancing/fill/spanners/rules, reverse wrapping, anonymous/generated items, `display: contents`, auto margins, effect-group pagination, resources, timeout, cancellation, limits, and diagnostics.

This checkpoint establishes the architecture and usable business-document output. It does not close the advanced-fidelity roadmap or justify full browser-equivalent HTML/CSS claims.

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

### Platform and resource envelope

- `OfficeIMO.Html` targets `netstandard2.0`, `net8.0`, and `net10.0`, plus `net472` on Windows. The executable benchmark harness targets `net8.0` and `net10.0`; end-to-end contracts also run on `net472`, and the existing trim/AOT smoke exercises the project-only dependency graph.
- Parsing, layout, Drawing projection, PNG encoding, SVG serialization, and PDF serialization do not require a browser, native rendering host, or a new runtime package.
- Installed fonts are inherently platform-specific. For byte- and metric-stable cross-platform output, callers should supply policy-approved TrueType `@font-face` data. Managed PDF system-font resolution uses a bounded 32-entry process cache and is activated only when scene text needs more than the standard WinAnsi path; restart the process after installing or removing fonts.
- Default source and DOM ceilings are 16,777,216 UTF-16 characters and 100,000 nodes. Resource, page, surface, layout-depth, track, column, gradient, shadow, and image-tile limits remain independently configurable.
- The non-packable `OfficeIMO.Html.Benchmarks` project defines review budgets for small and standard report classes across parse, styles, prepared layout, combined layout, Drawing, PNG, SVG, WinAnsi PDF, and multilingual PDF lanes. Allocation ceilings are stable review triggers; timing regressions are compared on the same machine instead of asserted by unit tests.

### Public API direction

The public names follow the shared OfficeIMO image-export conventions:

```csharp
HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
byte[] png = source.ToPng(options);
string svg = source.ToSvg(options);

IReadOnlyList<OfficeImageExportResult> pages =
    source.ExportImages(OfficeImageExportFormat.Png, pagedOptions);

PdfDocument pdf = source.ToPdfDocument(pdfOptions);
byte[] pdfBytes = source.ToPdf(pdfOptions);
source.SaveAsPdf("report.pdf", pdfOptions);
```

The image surface belongs to `OfficeIMO.Html`; the PDF extension surface remains in `OfficeIMO.Html.Pdf`. `To...` returns in-memory output, `To...Document` returns a typed model, and `SaveAs...` requires a file or stream destination. Async variants use the same names with `Async`, with cancellation and timeout carried through the whole operation.

## Required Feature Contract

### HTML and resources

- HTML5 parsing, malformed-markup recovery, base URI handling, and nested documents.
- Inline, embedded, policy-approved linked, and recursively imported stylesheets. TrueType font URLs and CSS background-image URLs are active; other CSS URL paint assets remain before this contract is complete.
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

Text shaping must be implemented with first-party managed code and existing OfficeIMO font primitives. The shared `OfficeIMO.Drawing.IOfficeTextShapingProvider` seam can remain useful, but this plan must not require a new HarfBuzz, Skia, or native-assets package.

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
- [x] Build a representative corpus: invoices, statements, reports, letters, certificates, catalog pages, dashboards, multilingual documents, and hostile-resource cases. The test-owned corpus covers every published `HtmlMarketScenarioCatalog` id so public examples cannot drift away from end-to-end evidence.
- [ ] Record expected geometry, page counts, extracted text, links, diagnostics, and visual baselines. The corpus now records mode, surface width, page count, minimum scene/headings, logical text, links, diagnostics, PNG signature, SVG text, and searchable-PDF readback for ten cases; approved pixel/byte visual baselines remain.
- [x] Define performance and memory budgets by document class. The non-packable benchmark project records small 10-row and standard 100-row stage budgets plus standard 40-row Drawing, PNG, SVG, WinAnsi PDF, and Unicode PDF output budgets; timing uses same-machine relative regression review rather than flaky wall-clock tests.

Exit gate: the public behavior and proof corpus exist before implementation details become permanent APIs.

### Phase 1 - Prepare the existing HTML foundation

- [x] `RenderAsync` carries caller cancellation through resource resolution, stylesheet/font preparation checkpoints, shared child-block construction, continuous placement, paged block loops, and forced-fragment loops. Image scene projection, per-page encoding boundaries, rendered-PDF font/page/visual projection, buffered serialization boundaries, and final async writes observe the same token.

- [x] Split `HtmlResourcePipeline.cs` into focused partials for element discovery, CSS discovery and syntax, custom-property resolution, selector matching, policy, and internal models.
- [x] Split `HtmlComputedStyleEngine.cs` by rules, media, supports, cascade, selector matching, CSS syntax, and internal models so the existing engine remains the shared owner.
- [x] Add cancellation and timeout propagation to direct-render resource loading and conversion orchestration.
- [x] Apply policy-approved linked stylesheets and recursive imports before page-rule and computed-style resolution, with cycle/depth/count/budget enforcement; nested background images and TrueType fonts are active while other CSS URL paint resources remain.
- [ ] Carry the same cancellation contract through any shared discovery/loading paths reused outside the direct renderer.
- [x] Add the `OfficeIMO.Drawing` project reference to `OfficeIMO.Html`.
- [x] Establish initial shared units, coordinate systems, colors, font descriptors, and diagnostic codes.

Exit gate: parsing, CSS, resources, and diagnostics can serve Word/RTF conversion and the new renderer without duplicated logic.

### Phase 2 - Computed values and formatting tree

- [ ] Complete cascade, inheritance, media, and computed-value handling required by the corpus. Custom-property inheritance/fallback/cycle handling, linked stylesheet order, and preserved font-family fallback lists are implemented.
- [ ] Build a typed formatting tree separate from the AngleSharp DOM.
- [ ] Resolve display roles, generated content, counters, replaced elements, pseudo-elements, and stacking contexts. Styled `::before`/`::after` text, attributes, scoped counters, cascade/specificity, inline/block-container placement, nested numeric stacking bands for absolute/fixed/relative/sticky block, root, flex, grid, and inline contexts, block/atomic-inline transform/opacity stacking contexts, and block/flex/grid/float/normal-inline replaced images are implemented; generated images, quote-state tokens, richer pseudo boxes, and normal-inline effect grouping remain.
- [ ] Preserve source and semantic references needed for diagnostics, links, accessibility, and extraction order. Visuals retain source descriptions, links, paint order, and layout ownership; heading fragments from the same source element retain one operation-scoped semantic identity and aggregate into ordered shared heading nodes, while paint-neutral groups preserve paragraph, heading, section/landmark, header/footer division, list, and table ownership. The shared render document retains bounded source title, language, and typed root direction metadata for all adapters. More specialized grouped DOM semantics and explicit reading-order nodes remain.

Exit gate: every rendered node has deterministic computed values or a stable unsupported-value diagnostic.

### Phase 3 - Typography and inline layout

- [ ] Consolidate font parsing, metrics, fallback, and glyph outlines in the existing shared owner, primarily `OfficeIMO.Drawing` where reusable. Scoped TrueType face registration, grapheme-safe deterministic measurement, non-BMP TrueType cmap lookup, CSS `@font-face` activation, raster/SVG/PDF reuse, rendered-PDF fallback wiring, authored CSS family-list preservation, and shared grapheme-safe family fallback planning from actual cmap coverage are implemented. HTML layout now resolves those runs once so PNG, SVG, and PDF consume the same selected family segments; WOFF/WOFF2/CFF decoding, local-source aliases, font subsetting, and richer OpenType metrics remain.
- [ ] Implement managed shaping, bidi resolution, ligatures, kerning, combining marks, and script-aware fallback. The computed-style cascade treats valid HTML `dir` attributes as low-specificity presentational hints, the shared document retains the resolved root direction, and rendered PDF maps RTL roots to right-to-left viewer page progression. Shared inline layout now positions simple RTL/Hebrew directional groups and numbers at grapheme-safe coordinates while retaining logical shared-scene order, and cmap-aware family fallback resolves once for all backends. `OfficeIMO.Drawing` supplies deterministic contextual presentation forms for the core Arabic alphabet before shared measurement and paint; paint-neutral logical-text groups retain the original source for `HtmlRenderDocument.Text`, heading aggregation, and PDF `ActualText`. Unsupported joining alphabets and explicit bidi embeddings/isolates retain separate stable diagnostics. The direct PDF adapter exposes the existing managed Latin-ligature mode and optional shaping-provider seam; broader OpenType shaping, mark positioning, and optional ligatures remain.
- [ ] Implement whitespace, line breaking, inline boxes, decorations, baseline alignment, and intrinsic text measurement. Logical `text-align:start/end` and the default start alignment now resolve against each element's computed direction while physical left/right remain physical; richer inline-box synthesis and typographic baseline work remain.
- [ ] Verify positioned glyphs round-trip to searchable/extractable PDF text. Current scalar text, Unicode fallback, explicit spaces, positioned Hebrew, and contextually shaped core Arabic round-trip exactly through logical `ActualText`; provider-shaped glyph sequences remain.

Exit gate: multilingual line boxes are deterministic across supported target frameworks and platforms.

### Phase 4 - Core layout

- [ ] Implement block flow, margin collapsing, padding, borders, backgrounds, sizing, and overflow. Background colors, normal-flow sibling/parent/empty margin collapse including negative margin sets and formatting boundaries, full per-corner circular/elliptical rounded color/border/gradient/URL boxes, independent side widths/colors and solid/dashed/dotted/double styles, matching offset outlines, layered outer/inset alpha/blur/spread box shadows, layered URL backgrounds with per-layer `auto`, `contain`, clipped `cover`, explicit size/position and bounded `repeat`/`space`/`round` patterns, opaque multi-stop linear and centered/off-center circular or elliptical radial gradients with percentage/length stops, duplicate hard edges, and CSS backward-position fix-up, root-canvas propagation, per-axis block/flex/grid/positioned overflow clipping, visual-box `overflow-clip-margin`, and document-root/body viewport overflow propagation are implemented. Interactive scrollbar UI and remaining gradient forms remain.
- [ ] Implement replaced elements and image/SVG intrinsic sizing. Raster image pixel/DPI dimensions; SVG absolute dimensions, partial-dimension/viewBox ratio derivation, and intrinsic aspect metadata; authored width/height; content-box/border-box sizing; bounded positive aspect ratios; min/max constraints; flex/float intrinsic planning; normal-inline wrapping and baseline participation; all five object-fit modes; common/edge-offset object positions; and per-corner circular/elliptical content clipping now share native placement/source-crop output. Common SVG primitives, paths, endpoint-form elliptical arcs, ordered nested transform attributes, bounded local shape/group/symbol references, bounded local object-bounding-box/user-space linear/radial fill and stroke servers, and positioned/affine searchable tspan text in `<img>` and CSS background sources now become one shared Drawing scene for PNG, SVG, and PDF; broader SVG features and replaced elements beyond images remain.
- [ ] Implement floats and relative/absolute/fixed/sticky positioning. Relative positioning is implemented for block and nested inline content with normal-flow-preserving fragmentation. Absolute elements resolve against initial, block, flex, grid-area, or measured wrapped-inline containing rectangles without consuming flow space or line width; automatic insets retain hypothetical block/inline flow markers; fixed elements repeat per page/viewport; numeric absolute/fixed/relative/sticky stacking contexts remain nested across block, root, flex, grid, and inline layout; and sticky elements produce stable in-flow document snapshots. Left/right and logical-side floats pack and wrap line bands with side-specific clearance inside inline formatting content; cross-block intrusion, `shape-outside`, and advanced float fragmentation remain.
- [ ] Complete table layout and row/cell fragmentation. Styled top/bottom captions, empty-table caption retention, bounded auto/fixed intrinsic text/replaced-image column sizing with column declarations and spanning constraints, separate-border spacing, collapsed-grid spacing suppression, one-pass table/column-group/column/row-group/row/cell border conflict resolution, occupancy-grid column placement, column spans, row-group-bounded row spans, span height, and span-safe row breaks are implemented; richer nested formatting-context intrinsic contributions and cell-content fragmentation remain.
- [ ] Implement flexbox, grid, and multicolumn formatting contexts. Row, column, and atomic inline flex cover ordering, main/cross reversal, wrapping and reverse wrapping, per-line grow/shrink/basis, min/max constraints, gaps, main/cross alignment, stretch, auto margins, definite-height vertical percentage bases, intrinsic cross-size planning, anonymous/generated items, `display: contents`, nesting, links, and safe single-axis boundaries. Block and atomic inline grid cover bounded fixed/percentage/auto/fr/minmax/integer/auto-fill/auto-fit tracks, implicit tracks, sparse/dense row and column placement, numeric/named spans, named areas, alignment, shared items/links, and span-aware row boundaries. Multicolumn layout covers bounded count/width resolution, normal and authored gaps, balance/auto fill, direct block spanners, common vector rule styles, generated-content order, overflow columns, and page-safe shared breaks. Full atomic baselines, subgrid, masonry, nested descendant spanners, advanced column-rule styles, and advanced fragmentation remain.

Exit gate: the continuous renderer can place the full target corpus without pagination.

### Phase 5 - Pagination and fragmentation

- [x] A single oversized vertical flex item, isolated non-spanning grid item, or lone non-spanning body table cell now projects nested safe block/line boundaries into its container flow, allowing page fragmentation without forced visual slicing; multi-line, overlapping flex/grid, and coordinated multi-cell row fragmentation remain advanced work.

- [ ] Complete page construction and page masters. Generic `@page` size/margins, print media selection, named page assignment, and generic/named first/left/right margin-content selection are implemented; named and pseudo-page body-geometry reflow remains.
- [ ] Complete break rules and continuation state. Stable nested text/child/table-row continuation, CSS `widows`/`orphans`, parity-correct left/right/recto/verso breaks, repeated leading `<thead>` and trailing `<tfoot>` groups, and row-span-safe boundaries are implemented; advanced fragmentation remains.
- [ ] Complete generated page content, counters, running headers/footers, and margin boxes. Quoted text, `counter(page)`, `counter(pages)`, and all sixteen standard margin boxes are implemented; running strings, element content, and richer generated content remain.
- [ ] Make backgrounds, borders, positioned elements, flex/grid content, and links fragment correctly. Relatively positioned visuals retain separate flow and paint coordinates so page ownership follows normal flow. Absolute visuals remain attached to their block/grid/inline containing rectangles, root absolute visuals appear once, fixed visuals repeat on every page without affecting fragmentation, negative/auto/positive absolute/fixed/relative/sticky stacking bands retain backend-neutral paint order, and effect groups rebase transforms while preserving isolated opacity across page fragments. Raster images, shared SVG drawings, image patterns, path-clipped image content, and all shared shapes that exceed one page retain their authored geometry and continue through vertically clipped page fragments across PNG, SVG, and PDF; gradients no longer restart per fragment, intermediate rounded corners are not invented, and line markers remain attached to authored endpoints. A supported single flex row that fits one page is kept together and moved to the next page, wrapped rows expose safe page boundaries only between complete flex lines, a single vertical flex column exposes boundaries between complete items, and grid exposes only row boundaries that no spanning item crosses; oversized line/item, multi-column flex, overlapping grid, inline box background/border fragment synthesis, and advanced grid fragmentation remain.

Exit gate: paged geometry and page counts match the accepted corpus.

### Phase 6 - Shared paint and semantic model

- [x] Convert the initial layout fragments into a backend-neutral ordered visual scene.
- [ ] Reuse or extend `OfficeIMO.Drawing` for paths, text outlines, images, gradients, transforms, clipping, shadows, and opacity. Shapes, uniform rounded rectangles, per-corner cubic rounded paths and clips, independent side paths, native dashed/dotted strokes, layered double strokes, offset outlines, layered outer/inset alpha/blur/spread shadows, text, intrinsic/aspect-sized images, object-fit/object-position source crops, bounded common SVG primitive and local paint-server scenes, active font faces, rectangular overflow clips, rounded/freeform path-clipped visual groups, bounded image-pattern visuals with independent repeat steps, multi-stop linear/radial gradient paint, nested affine effect groups, and isolated group opacity are shared now; advanced SVG features, filters, blend modes, and other advanced paint effects remain.
- [ ] Keep searchable text runs and semantic nodes beside visual operations. Searchable HTML text, positioned/affine SVG tspan runs, links, and retained heading/paragraph roles are implemented. Multi-fragment HTML headings share stable semantic identity and aggregate into a backend-neutral navigation node. Paint-neutral shared semantic groups preserve multi-run paragraph and H1-H6 ownership, section/landmark and header/footer divisions, list/item/label/body, and table/caption/row/header-cell/data-cell ownership, scope, and spans through translation, pagination, PNG/SVG traversal, font discovery, and typed PDF canvas containers. More specialized reading-order grouping remains.
- [ ] Complete deterministic z-order, clipping, resource reuse, and fallback diagnostics. Surface/root/content paint order, numeric negative/auto/positive absolute/fixed/relative/sticky stacking bands across block/flex/grid/inline contexts, block/atomic-inline transform and isolated-opacity contexts, nested positioned/effect contexts, backend-neutral nested overflow clips, back-to-front layered URL and linear/radial-gradient backgrounds, clipped cover, repeat patterns, encoded-image snapshot reuse, PDF image-XObject deduplication, native multi-stop axial/radial PDF gradient stitching, nested transparency Form XObjects, and stable layer/repeat/value/tile/stop-limit/positioning/overflow/effect diagnostics are implemented; normal-inline effects, blend modes, and remaining gradient forms remain.

Exit gate: a single rendered result contains everything needed by both image and PDF backends.

### Phase 7 - HTML to PNG and SVG

- [x] Use the shared `HtmlRenderOptions` contract for PNG and SVG output, aligned with `OfficeIMO.Drawing` export behavior without introducing a second HTML image-options type.
- [x] Add continuous `ToPng`, `ToSvg`, file, stream, synchronous, and asynchronous APIs.
- [x] Add paged `ExportImages` APIs with page numbering and diagnostics.
- [x] Carry active TrueType web fonts into PNG rasterization and embed them as data-backed `@font-face` definitions in SVG output.
- [ ] Complete maximum surface, tiling, scale, DPI, transparency, and background behavior. Surface limits, scale, and background are implemented.
- [ ] Activate image baselines for both paged and continuous modes.

Exit gate: HTML image output uses only `OfficeIMO.Html` plus existing OfficeIMO projects and produces no PDF intermediate.

### Phase 8 - Direct HTML to PDF

- [x] Add direct paged `ToPdf`, `ToPdfDocument`, `ToPdfResult`, and destination-only `SaveAsPdf` APIs in `OfficeIMO.Html.Pdf` over the shared rendered-document model.
- [ ] Complete mapping to PDF structures. Searchable Unicode HTML and positioned/affine SVG tspan text with managed fallback controls, exact explicit spaces, active TrueType web-font embedding, basic shapes, multi-stop axial/radial gradients through native PDF stitching functions, raster and grouped vector figures with alternative text, external links, affine effect transforms, isolated transparency Form XObjects, tagged document markers, shared Sect/Div plus single-owner multi-run H1-H6/P structure, nested native PDF outlines, shared L/LI/Lbl/LBody list hierarchy, and shared Table/Caption/TR/TH/TD hierarchy with scope and span attributes are implemented; more than three simultaneous distinct web-font families currently diagnose and fall back because the PDF writer exposes three semantic generated-font slots, while positioned shaped glyphs and more specialized reading-order semantics remain.
- [x] Keep HTML-to-Word, HTML-to-Markdown, HTML-to-Excel, and HTML-to-PowerPoint as explicit target projections in their owning packages instead of hiding them behind PDF profiles.
- [x] Add async/cancellable save APIs with explicitly buffered final PDF serialization.
- [ ] Validate page geometry, extraction, links, outlines, metadata, encryption, and tagged structure. Rendered PDF now carries retained HTML title and document language into PDF metadata/catalog, requests display of the document title, maps resolved RTL root direction to viewer page progression, preserves logical extraction for positioned RTL text through `ActualText`, emits typed H1-H6/P marked content, retains nested heading outlines with absolute page destinations, and proves raster plus multi-operation SVG image alt text through Figure structure; richer structure and encryption-profile coverage remain.

Exit gate: PDF and image output agree on layout while PDF preserves text and document semantics.

### Phase 9 - Advanced fidelity

- [ ] Complete advanced SVG, filters, masks, blend modes, and raster fallbacks. A bounded primitives, cubic-normalized path, affine transform-attribute, local shape/group/clipped-symbol reference, local object-bounding-box/user-space linear/radial paint-server with affine linear and axis-aligned radial transforms plus bounded linear repeat/reflect expansion, and positioned/affine/searchable tspan scene with glyph-scaled text lengths is now shared across `<img>` and CSS background sources in the current output backends, alongside layered inset/outer spread shadows; rotated/sheared or repeating radial paint servers, gradient/per-glyph/path text, spacing-only length adjustment, masks, filters, and CSS-style transforms remain.
- [ ] Complete difficult table/flex/grid/multicolumn fragmentation cases.
- [ ] Expand modern CSS value/functions and generated-content behavior.
- [ ] Add managed hyphenation dictionaries only if they can be shipped without a new runtime dependency and with acceptable package size.

Exit gate: every accepted feature has corpus proof and every unimplemented feature has a stable diagnostic.

### Phase 10 - Hardening and release readiness

- [ ] Add fuzz/hostile-input, resource-budget, timeout, cancellation, and deterministic-output tests. Byte-identical PNG, SVG, and rendered-PDF output for identical fully resolved input/options is now protected across supported target frameworks. Source strings are rejected before parsing above the configurable character budget, parsed DOMs are rejected before styling above the configurable node budget, and sync/async paths share the typed limit contract; broader hostile corpus and fuzz coverage remain.
- [x] Add an executable NativeAOT/trimming smoke that parses and lays out HTML, emits dependency-free SVG and PNG, creates rendered searchable PDF, reads the PDF marker back, and runs in the existing AOT/trim CI lane. The smoke has no package references; it consumes only the repository projects.
- [x] Benchmark parse, style, prepared layout, combined parse/style/layout, Drawing projection, PNG, SVG, and rendered PDF independently. The output lane covers both WinAnsi and multilingual Unicode text, uses the repository's existing BenchmarkDotNet development dependency, and adds nothing to shipped packages.
- [x] Add deterministic Markdown support-matrix generation directly from `HtmlConversionProfileContracts` and the ordered `HtmlDiagnosticCatalog`, including a no-BOM file writer so docs and release tooling can publish the same source of truth without a parallel hand-maintained capability list.
- [x] Document platform/target-framework differences and memory limits. The plan now records target frameworks, executable cross-target/AOT evidence, platform-font variability, process-cached managed fallback behavior, the no-browser/no-native-host boundary, explicit source/DOM and renderer budgets, and benchmark review ceilings by document class.

Exit gate: support claims are generated from passing evidence, not maintained as an aspirational list.

## Existing Issues to Close

These are part of the roadmap, but they should not distort the dependency order of the renderer.

| Issue | Required action | When it blocks |
|---|---|---|
| Legacy PDF encryption output | Add a modern standard-security writer, prefer AES-256, make modern encryption the default, and keep RC4 only as an explicit legacy option if required | Before calling direct HTML-to-PDF production-ready for sensitive documents; can be implemented in parallel with layout |
| Incomplete first-party complex text shaping | Extend the managed core-Arabic contextual-form path to explicit bidi controls, broader joining alphabets, mark positioning, and required OpenType behavior without native or external packages | Before Phase 3 exits and before broad multilingual fidelity claims |
| Buffered HTML-to-PDF orchestration | Async resource resolution and cancellation are implemented; keep final serialization explicitly documented as buffered until incremental PDF writing exists | Before claiming streaming output |
| Incomplete computed-style surface | Extend the existing style engine with typed computed values and diagnostics | Before layout phases can be correct |
| Partial end-to-end visual proof | The ten-case market corpus now proves paged/continuous geometry, logical text, links, diagnostics, PNG/SVG production, and searchable PDF readback; commit approved pixel/byte baselines next | Before premium fidelity claims or release |

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
