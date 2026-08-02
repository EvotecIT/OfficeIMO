# OfficeIMO roadmap

This is the repository's single product backlog. It contains open work only. Implemented behavior is documented in package READMEs, support matrices, generated inventories, and current-state guides linked from [the documentation index](README.md).

An item belongs here when it has a clear product outcome and an owning package. Implementation checkpoints, completed task lists, release-wait notes, and competitor parity tables do not belong here.

## Release-wide quality

- [ ] Generate a machine-readable capability manifest that distinguishes author, read, edit, preserve, inspect, convert, export, reject, and unsupported states across packages.
- [ ] Use that manifest to generate compatible sections of package READMEs, website capability pages, MCP discovery, and support matrices where one source can truthfully own the claim.
- [ ] Expand cross-producer fixture corpora with producer/version provenance and stable package or semantic diff policies.
- [ ] Keep correctness, file size, elapsed time, peak memory, allocation, cancellation, and deterministic-output evidence reproducible across macOS, Linux, and Windows.
- [ ] Add shared conversion reports and strict no-loss policies wherever an adapter can simplify, omit, rasterize, or preserve unsupported content.
- [ ] Keep public surfaces small: add reusable capability to the owning core package and keep CLI, PowerShell, website, MCP, Reader, and converter adapters thin.

## Excel

- [ ] Build one reusable reference syntax tree and rewriter for formulas, names, tables, charts, pivots, print definitions, and structural edits.
- [ ] Provide transactional row, column, and cell insertion/deletion plus copy, move, and transpose, with a dry-run mutation plan and post-edit package diagnostics.
- [ ] Complete AutoFilter criteria/state, table schema mutation, formula-aware search, range algebra, named styles, view/print state, conditional-formatting, and sparkline lifecycles.
- [ ] Add format-neutral public management for allowed-edit ranges and ignored-error regions, including preservation and remapping through structural edits.
- [ ] Complete A1/R1C1 conversion and explicit authored, cached, evaluated, dirty, deferred, and unsupported formula states, including dynamic-array metadata and high-value function clusters.
- [ ] Deepen pivot, slicer, timeline, modern-chart, query-backed source, and shared-cache workflows.
- [ ] Add native in-cell images and preserve their behavior through sorting, filtering, resizing, copying, and structural edits.
- [ ] Preserve additional relationship-backed drawings, workbook-level structures, charts, and template bindings when loading, cloning, editing, and saving complex imported workbooks.
- [ ] Add a memory-bounded edit path for large existing workbooks with configurable budgets and deterministic cancellation.
- [ ] Keep XLSX/XLSB/CSV performance evidence competitive on macOS, Linux, and Windows without platform-specific claims unsupported by the cross-platform matrix.

## PowerPoint and Visio

- [ ] Broaden PowerPoint editing, template-layout, chart/table, animation-preservation, and source-producer corpus coverage through the existing presentation API.
- [ ] Add fixture-backed PowerPoint authoring only where current demand justifies it, including shape/text/chart animations, broader SmartArt and custom geometry, advanced chart families, comment/reply mutation, OLE, macros, and custom shows.
- [ ] Establish repeatable PowerPoint create/save, open/edit/save, image-export, and PDF-export baselines for small, normal, and large decks before setting regression budgets.
- [ ] Improve Visio dense layout, label placement, routing, graph and data-driven diagrams, and editable existing-diagram workflows.
- [ ] Expand the Visio stencil platform while preserving package licensing and unsupported content explicitly.
- [ ] Prove Visio template, stencil, and macro-enabled package variants with representative load, reader, and conversion fixtures before advertising them beyond the current `.vsdx` boundary.
- [ ] Keep generated Visio galleries structurally validated and visually reviewed across representative diagram families.
- [ ] Harden optional desktop-reference validation without making Microsoft Office or Visio a runtime dependency.

## Word

- [ ] Deepen advanced drawing, imported chart mutation, SmartArt editing, grouping, anchoring, and Word-exact layout evidence beyond the current bounded shapes.
- [ ] Extend structured comparison and redline generation one documented shape at a time while keeping effective-formatting and true move semantics explicit until their stable contracts exist.
- [ ] Complete the remaining locale-sensitive, layout-dependent, complex-table, nested-instruction, and native `LISTNUM` field profiles without approximating unsupported results silently.
- [ ] Broaden Word/HTML reciprocal fidelity for ruby, forms, nested tables and list blocks, figures, comments, headers, footers, sections, properties, and accessibility metadata; finish mapped CSS-value diagnostics and aggregate budgets for supported non-image resources.
- [ ] Expand producer-provenanced review, redline, template, mail-merge, legacy DOC, Word/HTML, rendering, and performance corpora while preserving guarded loss reports and avoiding unsupported native DOC authoring claims.

## PDF, HTML, and image rendering

- [ ] Expand arbitrary-producer PDF rendering for Type 3/CFF fonts, ICC and currently unsupported color spaces, advanced patterns, and incomplete optional-content cases while keeping every fallback diagnosed and bounded.
- [ ] Deepen PDF annotation, form, tagged-structure, searchable-text, outline, metadata, encryption-profile, signature-validation, redaction-verification, and source-conversion evidence against cross-producer fixtures.
- [ ] Complete the remaining HTML cascade, shaping, bidi, inline-layout, intrinsic-sizing, table, flex, grid, multicolumn, pagination, page-master, fragmentation, and advanced SVG cases recorded as partial in the support matrix.
- [ ] Expand hostile-input, fuzz, aggregate-budget, timeout, cancellation, deterministic-output, and approved visual-baseline coverage across PDF, HTML, SVG, and raster paths at representative sizes, DPI values, fonts, and platforms.
- [ ] Keep `OfficeIMO.Drawing` as the reusable owner for codecs, placement, text layout, shapes, paths, colors, gradients, clipping, effects, and batch-export policy while format adapters remain thin.

### Image-export evidence

- [ ] Extend `OfficeIMO.Drawing` with reusable bounded codec, geometry, text-shaping, image, chart, streaming, cancellation, budget, and diagnostic contracts needed by more than one document package.
- [ ] Burn down `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, and `OfficeIMO.Word` visual-fidelity gaps with focused fixtures for worksheet objects and styling, slide inheritance and grouped content, and estimated Word pagination, overflow, and fallback reporting.
- [ ] Expand `OfficeIMO.Html`, `OfficeIMO.OneNote`, and `OfficeIMO.Visio` visual galleries across continuous and paged resources, real-world notebook content, and loaded diagrams while keeping allocation bounded before large working surfaces are created.
- [ ] Expand `OfficeIMO.Pdf` rendering evidence for operators, fonts, images, transparency, forms, annotations, and conservative arbitrary-producer coverage while retaining source-conversion warnings.
- [ ] Keep shared visual QA in `OfficeIMO.Shared.Tests`: approved PNG difference metrics, renderable/nonblank proof, stable diagnostics, and optional external reference tools that never become product dependencies.

## Markdown and text formats

- [ ] Close the remaining CommonMark 0.31.2 inventory failure and broaden GFM evidence without changing OfficeIMO-specific profile defaults.
- [ ] Finish generic-attribute ownership across the remaining supported block and inline families, including containers, source-backed edits, HTML output, and Markdown writing.
- [ ] Keep precise source locations partial until lossless trivia, delimiter tokens, original-to-normalized mapping, generated-node semantics, and broader source edits share one documented model.
- [ ] Make the semantic tree the canonical behavior owner and finish stable associations between semantic subobjects and source syntax.
- [ ] Complete lossless trivia and delimiter coverage, original-to-normalized source mapping, generated-node semantics, and broader source-edit/round-trip writing.
- [ ] Broaden parser, renderer, writer, transform, and extension contracts while keeping raw HTML grammar separate from security and host policy.
- [ ] Decide optional syntax ownership only when a real use case requires it, including grid tables, math, media, figures, diagrams, citations, footers, globalization, pragma lines, and container variants.
- [ ] Harden RTF untrusted-input limits, safe HTML output profiles, semantic-loss reporting, cancellation, structural editing, Word bridging, and performance baselines beyond the current tested profiles.
- [ ] Expand the provenance-recorded RTF producer corpus beyond current Word, Outlook, and LibreOffice evidence to Google Docs, macOS TextEdit/RTFD, EHR/CRM/helpdesk generators, and commercial libraries.
- [ ] Keep AsciiDoc and LaTeX support inside their documented bounded profiles; expand native syntax, semantic editing, adapters, and diagnostics only with source-preserving proof.
- [ ] Expand OpenDocument style, formula, drawing, embedded-content, signature, encryption, and producer-corpus coverage while preserving unknown package content.

## Reader and document intelligence

- [ ] Keep `OfficeIMO.Reader.Core` dependency-light while expanding the stable rich-result contract for pages, blocks, tables, links, forms, assets, visuals, OCR candidates, chunks, metadata, and source references.
- [ ] Deepen PDF logical-model projection, structured tables, assets, visual extraction, hierarchical chunks, and format-specific provenance.
- [ ] Define the deferred generic `ExtractStructured<T>()` contract only after the delivered non-generic structured extractor and processor pipeline have downstream compatibility evidence; keep model/client SDK dependencies outside Core.
- [ ] Keep OCR and other heavy/platform-specific providers optional; define their input, timeout, cancellation, and diagnostic contracts at the Core boundary.

## Email, stores, and cloud adapters

- [ ] Decide whether and when to add optional OneNote cloud transport. Any design must preserve native `.one` as the local-file goal, keep account/scope/revision/retry/data-loss policy outside `OfficeIMO.OneNote`, avoid implying that HTML projection is native `.one` fidelity, and not depend on GraphEssentialsX unless its eventual license is compatible with this repository.

## Browser and agent surfaces

- [ ] Establish measurable browser-converter bundle, startup, and memory budgets and close the remaining Unicode font-packaging diagnostics while retaining local-only processing and documented input limits.
- [ ] Keep the OfficeIMO CLI and STDIO MCP server bounded, query-first for stores, rooted to explicit file-system access, and backed by the same public OfficeIMO APIs.
- [ ] Share one conversion capability model across documentation, MCP discovery, and the browser UI.
- [ ] Keep agent skills and PowerShell commands as thin workflow surfaces over the owning packages.

## Completion rule

Remove an item when its public API, compatibility boundary, tests, generated evidence, and user documentation agree. GitHub Releases records delivered history, while `MIGRATION.md` retains only upgrade actions; this file does not retain completed milestones.
