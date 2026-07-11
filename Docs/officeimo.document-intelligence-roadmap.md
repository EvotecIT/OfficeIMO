# OfficeIMO Document Intelligence Roadmap

Date: 2026-07-10

## Purpose

This roadmap describes the next OfficeIMO document-intelligence layer: a dependency-light, deterministic, .NET-native stack for reading, converting, inspecting, exporting, and preparing documents for search, automation, reporting, and AI ingestion.

The roadmap is OfficeIMO-owned and intentionally product-general. The goal is to make the OfficeIMO family feel coherent:

- `OfficeIMO.Reader` owns ingestion, chunks, folder traversal, capability discovery, and host bootstrap.
- `OfficeIMO.Markdown` owns typed Markdown, Markdown writing, HTML rendering, profiles, and transforms.
- `OfficeIMO.Pdf` owns PDF creation, parsing, logical readback, extraction, manipulation, forms, compliance readiness, and first-party PDF export primitives.
- `OfficeIMO.Drawing` owns reusable visual intent such as colors, chart snapshots, text measurement, raster/vector primitives, image metadata, gradients, shadows, transforms, clipping, and managed font support.
- `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, `OfficeIMO.Visio`, `OfficeIMO.Word.Html`, `OfficeIMO.Markdown.Html`, and the PDF adapter packages own format-specific models and should feed shared ingestion/export contracts through narrow adapters.

## Source Documents

Keep this roadmap aligned with the current owner documents and proof manifests:

- `Docs/officeimo.reader.modular-roadmap.md`
- `Docs/officeimo.pdf.current-state.md`
- `Docs/pdf-conversion-scenarios.json`
- `Docs/officeimo.visio.roadmap.md`
- `Docs/officeimo.visio.premium-showcase.md`
- `Docs/officeimo.word-html-roadmap.md`
- `Docs/officeimo.word-html-support-matrix.md`

## Where We Are Now

### Reader

`OfficeIMO.Reader` is already a stable ingestion facade. It reads Word, Excel, PowerPoint, Markdown, PDF, HTML, ZIP, EPUB, CSV, JSON, XML, YAML, and plain text-like formats into `ReaderChunk` instances with stable IDs, location metadata, hashes, table metadata, folder traversal, warning chunks, progress callbacks, detailed source summaries, handler registration, capability manifests, and host bootstrap helpers.

`OfficeIMO.Reader.Pdf` registers PDF ingestion through the same reader facade. The shared `OfficeDocumentReadResult` model now carries chunks, assets, diagnostics, source maps, and normalized format-specific content from one extraction run. Further format fidelity belongs in the owning format packages and their narrow Reader adapters, not in another central reader switch statement.

Reader hosts can now use `OfficeDocumentReaderBuilder` to compose modular adapters and custom handlers, then freeze that configuration into an immutable `OfficeDocumentReader`. Each instance owns its routing snapshot, so concurrent services can use different handlers for the same extension without process-global interference. The static registration API remains a compatibility surface for existing applications.

### Markdown

`OfficeIMO.Markdown` provides a no-runtime-dependency typed Markdown object model, parser, renderer, writer profiles, HTML rendering, syntax metadata, transforms, portable profiles, semantic fenced-block extension points, table support, images, links, front matter, footnotes, callouts, and host-oriented rendering paths.

`OfficeIMO.Markdown.Pdf` is now part of the conversion surface, so Markdown can participate directly in the first-party PDF verification loop. Markdown should remain the canonical portable text model for readable output, search indexing, documentation export, and AI-friendly content streams.

### HTML And Word HTML

The Word/HTML surface has advanced enough to be treated as a first-class document-intelligence path rather than a side converter. `OfficeIMO.Word.Html`, `OfficeIMO.Markdown.Html`, and `OfficeIMO.Html.Pdf` now cover practical HTML flows with better table import, CSS diagnostics, semantic profiles, and PDF output.

The next HTML work is to align Word, Markdown, PDF logical readback, and Reader chunks around the same block/table/asset diagnostics instead of treating HTML as a separate endpoint.

### Word Market Readiness

`OfficeIMO.Word` keeps its non-PDF market-readiness story grounded in four workflows:

- template and document assembly polish
- review, redline, and diff workflows
- HTML and Markdown conversion fidelity with proof artifacts
- real-document docs and showcase output

Word-to-PDF parity remains owned by the `OfficeIMO.Word.Pdf` / `OfficeIMO.Pdf` workstream; the Word market-readiness story should not make PDF fidelity promises.

### Drawing

`OfficeIMO.Drawing` is becoming the shared visual primitive layer used by PDF, Excel/PDF, PowerPoint/PDF, and Visio export work. It now supports reusable chart snapshots, chart rendering quality reports, text measurement, managed TrueType/OpenType font paths, color and image abstractions, vector/raster primitives, and drawing quality diagnostics.

This is important for document intelligence because visual content needs a shared language before it can be rendered, inspected, compared, described, or exported consistently.

### PDF

`OfficeIMO.Pdf` is now a serious first-party document engine, with `Docs/officeimo.pdf.current-state.md` as the canonical state file:

- Fluent PDF creation with headings, paragraphs, rich text, lists, panels, rows/columns, tables, images, drawing primitives, links, bookmarks, headers, footers, page setup, themes, standard-font measurement, viewer/catalog options, and visual baselines.
- PDF reading with `PdfReadDocument`, `PdfInspector`, `PdfValidator`, `PdfTextExtractor`, `PdfLogicalDocument`, and logical Markdown options.
- Logical readback for pages, text blocks, headings, paragraphs, list items, tables, images, links, form widgets, metadata, outlines, page labels, named destinations, open actions, viewer preferences, and simple AcroForm fields.
- Markdown extraction through `PdfTextExtractor.ExtractMarkdown(...)`, page/range variants, output stream/path helpers, and `PdfLogicalDocument.ToMarkdown(...)`.
- Text, structured text, table, and image extraction with page-range support and deterministic output-file naming.
- Page manipulation: split, extract, merge, import, duplicate, move, delete, reorder, rotate, stamp, watermark, and metadata editing.
- Forms: inspect simple fields, create simple fields, fill simple values, and flatten supported text/choice/button widgets.
- Compliance-readiness proof plumbing, a manual strict-validator workflow, and visual review gallery support.

The first-party adapter family is now broader:

- `OfficeIMO.Word.Pdf`
- `OfficeIMO.Excel.Pdf`
- `OfficeIMO.PowerPoint.Pdf`
- `OfficeIMO.Markdown.Pdf`
- `OfficeIMO.Html.Pdf`
- `OfficeIMO.Reader.Pdf`

`Docs/pdf-conversion-scenarios.json` makes conversion coverage observable across Word, Excel, Markdown, HTML, and PowerPoint. The remaining PDF work is fidelity, coverage, corpus breadth, structured output quality, and stronger diagnostics, not proving the foundation exists.

### Visio

`OfficeIMO.Visio` is now a broad dependency-light VSDX authoring, editing, inspection, export, and proof surface:

- VSDX create/load/edit/save, validation, unknown-content preservation, pages, shapes, connectors, groups, layers, hyperlinks, User cells, Shape Data, protection, page settings, backgrounds, masters, stencils, styles, comments, containers, swimlanes, and package-backed masters.
- High-level builders for flowcharts, block diagrams, dependency diagrams, architecture diagrams, networks, topology, swimlanes, org charts, timelines, sequences, generic graphs, and reusable gallery/showcase scenarios.
- Built-in and package-backed stencil catalogs with searchable metadata, preview payloads, learned dimensions, connection points, and typed stencil migration maps.
- Routing, label placement, polish passes, visual quality analysis, inspection snapshots, stencil profiles, generated galleries, premium gallery baselines, showcase summary artifacts, native SVG export, managed native PNG export, and optional desktop Visio validation/export helpers.

For document intelligence, Visio should become another structured source: extract pages, diagram nodes, connectors, labels, Shape Data, hyperlinks, stencil identity, inspection snapshots, visual quality summaries, and optional SVG/PNG previews into the shared model.

### Office Exporters

First-party PDF export is now a platform direction, not a single package:

- Word-to-PDF covers common document sections, page setup, headings, paragraphs, lists, links/bookmarks, TOC entries, tables, images, content controls, simple forms, headers/footers, footnotes/endnotes, metadata, and warnings.
- Excel-to-PDF covers visible sheets, print areas, worksheet tables, merged cells, images, charts through shared drawing snapshots, headers/footers, page setup, links, styles, number formats, and warnings.
- PowerPoint-to-PDF covers page-sized slide canvases, shapes, text, images, tables, charts through shared drawing snapshots, backgrounds, and diagnostics.
- Markdown-to-PDF and HTML-to-PDF bring portable document profiles into the same engine.

These exporters should feed back into the same verification, logical readback, visual proof, compliance-readiness, and asset pipeline used by ingestion.

## Target Architecture

The next layer should be a shared document readback model rather than a new one-off parser.

Proposed model names can evolve, but the shape should be stable:

```text
OfficeDocumentReadResult
  Source
  CapabilitiesUsed
  Markdown
  Html
  Json
  Chunks
  Assets
  Diagnostics
  Pages
  Blocks
  Tables
  Images
  Links
  Forms
  Visuals
  SourceMap
```

Core principles:

- One extraction run should be able to produce chunks, Markdown, JSON, HTML, assets, and diagnostics.
- Format-specific packages keep ownership of their real models.
- The shared model stores normalized intent, not raw PDF operators, Open XML parts, VSDX ShapeSheet details, or HTML parser internals.
- Unsupported content should produce diagnostics and preserve source references where possible.
- Heavy or platform-specific work remains optional and isolated.

## Current Branch Implementation

This branch starts the shared model and adapter path:

- `OfficeDocumentReadResult` and deterministic JSON serialization live in `OfficeIMO.Reader`.
- `DocumentReader` can return the shared envelope for existing chunk-based ingestion without changing `DocumentReader.Read(...)`, including generic summary metadata for chunks, blocks, tables, visuals, and known source containers.
- Excel table readback now preserves workbook path, sheet, A1 range, source chunk, and table index metadata so sheet containers can own their tables in the shared result.
- Markdown table readback now preserves source and normalized line spans, heading context, block anchors, block kind, and stable table indexes in the shared result.
- Markdown visual fenced blocks now preserve source and normalized line spans, heading context, block anchors, block kind, payload hash, and JSON location metadata in the shared result.
- Visual-only facades now exist for the core Reader pipeline, returning normalized visual payloads with source locations without requiring callers to build the full read-result envelope.
- Word, Excel, and PowerPoint read-result paths now populate the shared asset manifest for embedded OpenXML image parts, including deterministic asset IDs, media type, extension, suggested filename, source relationship identity, intrinsic pixel dimensions for PNG/JPEG/GIF/BMP payloads, Excel image alt text where present, payload hash, payload bytes for materialization, sheet-level placement for Excel, and slide-level placement for PowerPoint.
- Asset-only facades now exist for the core Reader pipeline, returning normalized asset manifests without requiring callers to inspect the full read-result envelope.
- `OfficeIMO.Reader.Pdf` maps logical PDF readback into the shared envelope and JSON output.
- `OfficeIMO.Reader.Visio` is an optional adapter over `OfficeIMO.Visio`, with page chunks, Shape Data tables, blocks, links, and optional SVG/PNG preview asset metadata.
- Reader handlers can register native path and stream `OfficeDocumentReadResult` delegates. The generic `DocumentReader.ReadDocument(...)` entry point preserves PDF and Visio rich results, while legacy chunk reads project the same result's chunks when a handler is rich-result-only.
- `OfficeDocumentReaderBuilder` and `OfficeDocumentReader` provide instance-scoped handler routing, capability manifests, file/stream/byte reads, and folder ingestion. Every modular Reader adapter exposes a matching `Add...Handler()` builder extension, while the static registry remains backward compatible.
- Reader registrations can expose native asynchronous rich-result delegates. `ReadAsync(...)` and `ReadDocumentAsync(...)` await those delegates directly, non-seekable inputs use an asynchronous bounded snapshot, and synchronous format engines use a bounded worker fallback. `ReadDocumentsAsync(...)` adds deterministic multi-file execution with explicit concurrency and document-count limits.
- Reader detection now reports extension kind, content kind, confidence, media type, bounded evidence, and mismatch state. Reads preserve known-extension behavior by default, can prefer content for mislabeled inputs, and can route unknown extensions to a unique registered handler by detected kind. Generic and native rich results expose detection, parsing, limit, truncation, unsupported-content, read, and OCR findings through structured diagnostics instead of warning strings alone.
- Reader instances can freeze an ordered processor pipeline with explicit throw, continue-with-diagnostic, or stop-with-diagnostic failure policy. Sync and async document, chunk, JSON, structured, and bounded batch reads use the same configured pipeline; folder chunk enumeration remains unchanged.
- Opt-in shared-model processors now normalize blocks, list and heading levels, tables, and links; classify repeated page-boundary artifacts; and filter assets together with dependent OCR candidates. Hosts can add typed sync or async processors without adding format-specific behavior to the facade.
- Bounded structured extraction now exposes metadata, forms, key/value and Path/Type/Value rows, Visio Shape Data, heading sections, named tables, chart summaries, quality/readiness summaries, and source diagnostics through a deterministic non-generic result and JSON serializer.
- Token-aware hierarchical chunking now keeps `ReaderChunk` as the leaf contract while adding document, page/slide/sheet, and heading nodes; exact source character spans; deterministic IDs/hashes; configurable overlap and context; host token counters; structured bounds diagnostics; and an independently versioned JSON sidecar.
- Document-level metadata entries now carry stable catalog, outline, destination, open-action, viewer-preference, and form-summary facts without making the shared Reader model depend on PDF-specific types.
- PDF source preflight capability flags now flow into metadata, and read/rewrite blockers flow into shared diagnostics as stable `pdf-read-blocker` and `pdf-rewrite-blocker` entries for file and stream readback.
- OCR readiness is represented as `OcrCandidates` plus `ocr-needed` diagnostics for image-only PDF pages and embedded Office image assets, without adding an OCR engine or service dependency to the core.
- Asset records include deterministic suggested filenames so hosts can write or index extracted images and previews without inventing adapter-specific naming rules.
- Materializable asset payloads can be written to a directory or streamed through caller-owned callbacks while JSON output remains manifest-only.
- Small materializable assets can be exposed as bounded, opt-in data URIs for HTML, sidecar JSON, and preview workflows.
- Asset materialization can opt into SHA-256 payload-hash validation before writing or streaming extracted assets.
- `ReaderTable` instances can be exported to deterministic CSV, Markdown, or JSON without format-specific adapter code.
- Table-only facades now exist for the core Reader pipeline, the PDF logical adapter, and the Visio inspection adapter.
- Table export bundles now package each `ReaderTable` with deterministic IDs, file-name stems, and CSV/Markdown/JSON sidecar payloads.
- Table export bundles can be written to a directory or streamed through caller-owned callbacks as deterministic `.csv`, `.md`, and `.json` sidecars.
- Visual export bundles now package each `ReaderVisual` with deterministic IDs, file-name stems, raw source payloads, payload extensions, and JSON sidecar payloads.
- Visual export bundles can be written to a directory or streamed through caller-owned callbacks as deterministic visual payload and `.json` sidecars.
- Focused tests cover the shared envelope for Markdown, Word, Excel, PowerPoint, PDF, and Visio across `net472`, `net8.0`, and `net10.0`.

## Output Contracts

### Markdown

Use `OfficeIMO.Markdown` as the final Markdown writer where possible. PDF logical Markdown already exists; the next step is to align Reader chunks, PDF logical output, Word/Excel/PowerPoint/Visio readback, and Markdown profiles so hosts can choose portable output consistently.

### JSON

Schema version 5 is the first stable JSON envelope for pages, blocks, tables, links, forms, diagnostics, assets, visuals, OCR candidates, chunks, metadata, and source references. The schema is embedded in `OfficeIMO.Reader`, packed as a versioned artifact, and guarded by strict deserialization and transport round-trip tests. Future breaking transport changes require a new schema version and explicit compatibility policy.

### HTML

Render through `OfficeIMO.Markdown` for portable document output. Use direct HTML when the source model has layout, review, or preview data that Markdown cannot express without losing useful structure. Align `OfficeIMO.Word.Html`, `OfficeIMO.Markdown.Html`, `OfficeIMO.Html.Pdf`, and PDF positioned-review HTML diagnostics around the shared block/source model.

### Chunks

Keep `ReaderChunk` as the ingestion contract. Token-aware hierarchy is an opt-in sidecar around those leaves, so stable v5 transport does not need another chunk model or breaking fields. Continue improving source maps and optional block/table/image/form/diagram references in owning adapters.

### Assets

Assets should have stable IDs, hashes, media type, source location, dimensions where known, and deterministic output filenames. PDF image extraction and Visio native preview proof point the way; extend the asset manifest across Office documents, HTML/EPUB, and Visio previews.

## Roadmap

### P0 - Keep Planning Current

Goal: make the roadmap OfficeIMO-owned, current, and easy to hand off.

- Keep this file as the general document-intelligence roadmap.
- Link it from adjacent docs only where useful; avoid duplicate dated snapshots.
- Keep the source-of-truth docs split by ownership: Reader modular roadmap, PDF current state, PDF conversion scenarios, Visio roadmap/showcase, and Word/HTML roadmap/support matrix.
- Avoid external-product-oriented naming in docs, tests, branches, and PR metadata.
- Re-run naming hygiene checks whenever this file is refreshed.

### P1 - Shared Read Result

Goal: one result envelope for all output forms.

- Add a read result model with Markdown, HTML, JSON, chunks, metadata, assets, diagnostics, pages, visuals, and source maps.
- Add capability metadata similar to `DocumentReader.GetCapabilityManifest()`.
- Start as internal or experimental until the model survives PDF, Word, Excel, PowerPoint, Markdown, HTML, and Visio examples.
- Provide convenience APIs over `DocumentReader`, for example `ReadDocument(...)`, `ReadDocumentJson(...)`, `ReadAsMarkdown(...)`, and `ReadAsHtml(...)`.
- Keep `DocumentReader.Read(...)` behavior stable.

### P2 - PDF Logical Model Integration

Goal: make current PDF logical readback the first full adapter into the shared model.

- Map `PdfLogicalDocument` pages, text blocks, headings, paragraphs, list items, tables, images, links, form widgets, metadata, outlines, destinations, viewer/catalog data, and diagnostics into the shared result.
- Reuse `PdfLogicalDocument.ToMarkdown(...)` rather than creating a second PDF Markdown path.
- Add JSON serialization for logical PDF readback.
- Add asset manifest output for PDF images and form/widget metadata.
- Add page-range readback APIs that preserve source page numbers and caller order.
- Keep `PdfInspector.Preflight(...)`, `PdfConversionReport`, and compliance-readiness diagnostics visible in the read result.

### P3 - Tables And Structured Blocks

Goal: make table and block output reliable enough for automation.

- Normalize table output across PDF, Excel, Word, Markdown, HTML, CSV, JSON, XML, YAML, and PowerPoint table adapters.
- Preserve table source references, row/column counts, span data, truncation state, captions/titles where known, and cell text.
- Promote existing PDF-to-Word/Excel table import helpers into the shared table contract once confidence metadata is available.
- Extend table-only extraction mode from Reader/PDF/Visio into examples, docs, and any remaining format-specific adapters.
- Improve PDF table detection with span geometry, repeated x-bands, ruling lines where available, and page-continuation heuristics.
- Expand deterministic CSV/Markdown/JSON table sidecar helpers into host-facing docs and end-to-end examples.

### P4 - Assets And Visuals

Goal: treat images, previews, charts, and diagram visuals as first-class output.

- Generalize the PDF image manifest into `DocumentAsset`.
- Continue enriching Word/Excel/PowerPoint image manifests with drawing anchors, crop metadata, broader alt/title metadata, and richer header/footer coverage.
- Add chart snapshots and drawing quality diagnostics from `OfficeIMO.Drawing`.
- Add Visio SVG/PNG preview output as visual assets when requested.
- Include asset hash, media type, extension, dimensions, source page/slide/sheet/diagram page, relationship or object identity where known, and warnings.
- Support output modes: asset-only readback, manifest-only envelope output, write-to-directory, embedded data URI for small assets, and caller-owned stream callbacks.

### P5 - Office Export Feedback Loop

Goal: use first-party PDF export as both product feature and verification path.

- Continue Word-to-PDF fidelity using `OfficeIMO.Pdf` primitives and warnings.
- Continue Excel-to-PDF coverage for print layout, charts, images, links, merged cells, styles, headers/footers, and diagnostics.
- Continue PowerPoint-to-PDF coverage for slide layout, shape/text/image/table/chart fidelity, grouped transforms, backgrounds, and diagnostics.
- Keep Markdown-to-PDF and HTML-to-PDF aligned with portable document profiles.
- After each export, use `PdfLogicalDocument`, `PdfInspector`, conversion scenarios, compliance proof packs, and raster/visual baselines to verify that generated output is readable and visually stable.

### P6 - Visio Reader And Diagram Intelligence Adapter

Goal: make VSDX useful to ingestion and automation hosts.

- Keep `OfficeIMO.Reader.Visio` as the optional adapter package unless package-boundary review says it belongs elsewhere.
- Extract diagram pages as chunks with title, shapes, connectors, labels, Shape Data, hyperlinks, stencil identity, inspection summary, visual quality summary, and preview asset references.
- Expose diagram tables/records where Shape Data schemas or graph records are present.
- Optionally emit SVG/PNG previews through `ToSvg(...)`, `SaveAsSvg(...)`, `ToPng(...)`, and `SaveAsPng(...)`.
- Keep VSDX semantics in `OfficeIMO.Visio`; Reader should adapt, not duplicate.

### P7 - HTML And Portable Document Bridge

Goal: align HTML, Markdown, Word, PDF, and Reader output into one portable document story.

- Map Word HTML import/export diagnostics into shared diagnostics.
- Reuse `OfficeIMO.Markdown` and `OfficeIMO.Markdown.Html` where semantic Markdown is the better intermediate.
- Keep `OfficeIMO.Html.Pdf` focused on semantic/document and positioned-review profiles rather than becoming a second HTML engine.
- Add fixtures that compare Word -> HTML -> PDF, Markdown -> HTML -> PDF, PDF -> HTML, and Reader HTML ingestion at the block/table/asset level.

### P8 - Processor Pipeline

Goal: deterministic cleanup and enrichment across formats.

- The ordered immutable pipeline, sync/async processor contract, step evidence, cancellation, and explicit failure policies are implemented.
- Opt-in processors cover header/footer/artifact classification, table cleanup, list and heading normalization, link normalization, and asset/OCR-candidate filtering.
- Host custom processors are supported through a typed interface, base class, or delegates. Instances used by a shared reader must be safe for concurrent calls.
- Form/key-value projection is available through structured extraction rather than a mutating cleanup step.
- Diagram shape/connector enrichment remains adapter-owned work for the Visio fidelity track.

### P9 - Structured Extraction

Goal: expose schema-friendly extraction without making the core depend on AI services.

- The bounded non-generic extractor covers key/value and Path/Type/Value rows, headings and following paragraphs, named tables, forms, Visio Shape Data, metadata, chart summaries, visual/table quality, chunk readiness, and security/OCR/limit diagnostics.
- Schema-friendly JSON serialization is available with an independent schema id/version. The stable rich transport remains `OfficeDocumentReadResult` version 5.
- `ExtractStructured<T>()` remains deferred until the non-generic contract has downstream use and compatibility evidence.
- Optional host-supplied model interfaces remain later work; cloud and client SDKs stay out of the core package.

### P10 - OCR-Ready Core

Goal: prepare the model and pipeline for OCR while keeping the core honest.

- Detect likely scanned pages or image-only regions. The first slice flags image-only PDF pages.
- Add `ocr-needed` diagnostics and `OcrCandidate` records.
- Add an `IOfficeOcrEngine` abstraction with line, word, character, bounding-box, confidence, language, and diagnostic output.
- Add merge rules that combine OCR output with native text spans without losing source coordinates.
- Treat OCR results as another source layer that can be compared with native text, not as a replacement for native text.
- Keep built-in high-quality OCR out of the no-dependency core unless a separate product decision is made.

### P11 - Optional Heavy Adapters

Goal: make advanced ingestion possible without changing the core dependency story.

Optional packages can be added after the core contracts are stable:

- OCR via Windows inbox APIs.
- OCR via configured external process or service.
- OCR via native engine package.
- Model-assisted structured extraction through host-provided clients.
- Audio/video/transcription adapters if needed by downstream products.

No optional adapter should be pulled transitively by `OfficeIMO.Reader`.

### P12 - Quality Gates And Evidence

Goal: prove document intelligence with tests users can trust.

- Golden Markdown snapshots.
- JSON schema snapshots.
- Stable chunk/source hash tests.
- Asset manifest snapshots.
- PDF logical readback tests.
- PDF conversion scenario tests.
- PDF compliance-readiness proof packs and strict-validator handoff artifacts.
- PDF visual review gallery artifacts.
- PDF raster baselines for visual output.
- Visio SVG/PNG/inspection/stencil-profile/visual-quality baselines.
- Visio showcase summary validation with proof/evidence totals and gallery artifacts.
- Word/HTML support matrix fixtures.
- Folder-ingestion summary tests.
- Cross-target builds for `net472`, `netstandard2.0`, `net8.0`, and `net10.0` where applicable.
- Reader-wide BenchmarkDotNet coverage for extraction, detection, transport, and parser/chunker isolation, with environment-qualified baseline notes for release decisions.

## Near-Term Implementation Slices

1. Keep this roadmap current and link it only from the active owner docs.
2. Evolve the stable `OfficeDocumentReadResult` model additively; use a new schema version for breaking transport changes.
3. Deepen PDF logical-document read-result coverage for richer image/form metadata, compliance diagnostics, destinations, outlines, and catalog evidence.
4. Align read-result Markdown through existing PDF logical Markdown and `OfficeIMO.Markdown`.
5. Keep deterministic schema version 5 JSON stable and expand nested schema detail as shared models mature.
6. Extend asset manifests to richer PDF form/widget assets, HTML/EPUB referenced media, and Office drawing anchors.
7. Expand table-only extraction examples and adapter coverage beyond the initial Reader/PDF/Visio facades.
8. Extend scan/OCR-needed diagnostics beyond image-only PDF pages into richer image-region and Office-document heuristics, still without an OCR implementation.
9. Extend the Visio adapter with stencil profile, inspection summary, visual quality summary, and optional write-to-directory previews.
10. Add HTML/Word/Markdown/PDF bridge fixtures at the block/table/asset level.
11. Add docs/tests showing Word, Excel, PowerPoint, PDF, Markdown, HTML, and Visio flowing into the same result shape.

## Non-Goals For The Core

- Built-in high-quality OCR in the base package.
- Mandatory cloud services.
- Mandatory native dependencies.
- Full fidelity for every possible PDF, Office, HTML, or VSDX feature in the first shared model.
- Silent best-effort readback that hides unsupported content.
- Editable Office package reconstruction from arbitrary PDF until logical readback, tables, coordinates, images, forms, and diagnostics are stable.
- Legacy binary Office formats unless a dependency-light parser becomes realistic.

## Design Rules

- Keep ownership clean: format packages parse their own formats.
- Keep wrappers thin over reusable .NET APIs.
- Prefer stable models and diagnostics over clever text-only projections.
- Preserve source references whenever possible.
- Treat visual proof, logical readback, compliance-readiness, and editable package fidelity as separate gates.
- Keep optional heavy features out of transitive core package paths.
- Let OfficeIMO outputs become OfficeIMO inputs; export, readback, and verification should reinforce each other.
