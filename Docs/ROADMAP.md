# OfficeIMO roadmap

This is the repository's single product backlog. It contains open work only. Implemented behavior is documented in package READMEs, support matrices, generated inventories, and current-state guides linked from [the documentation index](README.md).

An item belongs here when it has a clear product outcome and an owning package. Implementation checkpoints, completed task lists, release-wait notes, architectural rules, and competitor parity tables do not belong here.

## Release-wide quality

- [ ] Extend the generated Office compatibility catalog beyond the current Word, Excel, and PowerPoint legacy-format families into a package-neutral operation model for create, read, edit, preserve, inspect, convert, export, reject, and unsupported behavior.
- [ ] Generate compatible package README sections, website capability pages, MCP discovery, and support matrices from that model wherever one source can truthfully own the claim.
- [ ] Expand cross-producer fixture corpora with producer/version provenance and stable package or semantic diff policies.
- [ ] Add reproducible correctness, file-size, elapsed-time, peak-memory, allocation, cancellation, and deterministic-output evidence for representative workloads on every supported operating system.
- [ ] Add shared conversion reports and strict no-loss policies wherever an adapter can simplify, omit, rasterize, or preserve unsupported content.

## Security and protected content

- [ ] Add interoperable ODF encryption/decryption only after an external producer corpus, explicit password and key policy, and fail-safe preservation evidence are available.

## PowerPoint and Visio

- [ ] Broaden imported-chart editing and image/PDF export fidelity across advanced chart families, with explicit diagnostics for producer-specific content that remains preservation-only.
- [ ] Expand editable SmartArt coverage only where imported layout and connection topology can be represented without changing diagram meaning; preserve or reject unsupported topologies explicitly.
- [ ] Extend typed animation timeline authoring and editing beyond the currently supported shape, text, and chart effects while preserving unsupported sequences losslessly.
- [ ] Deepen typed media and OLE editing for linked and embedded audio, video, playback metadata, and package payloads while continuing to preserve unsupported content.
- [ ] Extend typed Visio editing for advanced nested containers, deeper swimlane metadata and automatic lane assignment, and richer threaded comment and author workflows.
- [ ] Complete advanced Visio resize-to-content and broader whole-diagram relayout and polish for dense imported diagrams.
- [ ] Add typed Visio APIs for data graphics, legends, and additional high-value ShapeSheet sections and formulas while continuing to preserve unmodeled package content explicitly.
- [ ] Prove Visio template, stencil, and macro-enabled package variants with representative load, reader, and conversion fixtures before advertising them beyond the current `.vsdx` boundary.

## Word

- [ ] Deepen advanced drawing, imported chart mutation, SmartArt editing, grouping, anchoring, and Word-exact layout evidence beyond the current bounded shapes.
- [ ] Extend structured comparison and redline generation one documented shape at a time while keeping effective-formatting and true move semantics explicit until their stable contracts exist.
- [ ] Complete the remaining locale-sensitive, layout-dependent, complex-table, nested-instruction, and native `LISTNUM` field profiles without approximating unsupported results silently.
- [ ] Broaden Word/HTML reciprocal fidelity for ruby, forms, nested tables and list blocks, figures, comments, headers, footers, sections, properties, and accessibility metadata; finish mapped CSS-value diagnostics and aggregate budgets for supported non-image resources.
- [ ] Expand producer-provenanced review, redline, template, mail-merge, legacy DOC, Word/HTML, rendering, and performance corpora while preserving guarded loss reports and avoiding unsupported native DOC authoring claims.

## PDF, HTML, and image rendering

- [ ] Expand arbitrary-producer PDF rendering for Type 3/CFF fonts, ICC and currently unsupported color spaces, advanced patterns, and incomplete optional-content cases while keeping every fallback diagnosed and bounded.
- [ ] Deepen PDF annotation, form, tagged-structure, searchable-text, outline, metadata, encryption-profile, signature-validation, redaction-verification, and source-conversion evidence against cross-producer fixtures.
- [ ] Complete the HTML cascade, shaping, bidi, intrinsic-sizing, table, flex, grid, multicolumn, pagination, page-master, fragmentation, advanced SVG, and validator-backed accessibility cases recorded as partial in the generated support matrix.
- [ ] Expand hostile-input, fuzz, aggregate-budget, timeout, cancellation, deterministic-output, and approved visual-baseline coverage across PDF, HTML, SVG, and raster paths at representative sizes, DPI values, fonts, and platforms.

### Image-export evidence

- [ ] Extend `OfficeIMO.Drawing` with reusable bounded codec, geometry, text-shaping, image, chart, streaming, cancellation, budget, and diagnostic contracts needed by more than one document package.
- [ ] Burn down `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, and `OfficeIMO.Word` visual-fidelity gaps with focused fixtures for worksheet objects and styling, slide inheritance and grouped content, and estimated Word pagination, overflow, and fallback reporting.
- [ ] Expand `OfficeIMO.Html`, `OfficeIMO.OneNote`, and `OfficeIMO.Visio` visual galleries across continuous and paged resources, real-world notebook content, and loaded diagrams while keeping allocation bounded before large working surfaces are created.
- [ ] Expand `OfficeIMO.Pdf` arbitrary-producer visual evidence for operators, fonts, images, transparency, forms, and annotations while retaining source-conversion warnings.

## Markdown and text formats

- [ ] Close the remaining CommonMark 0.31.2 inventory failure and broaden GFM evidence without changing OfficeIMO-specific profile defaults.
- [ ] Finish generic-attribute ownership across the remaining supported block and inline families, including containers, source-backed edits, HTML output, and Markdown writing.
- [ ] Complete trivia and delimiter coverage, original-to-normalized mapping, generated-node semantics, and stable associations between the semantic model and source syntax before claiming precise locations or lossless arbitrary edits.
- [ ] Harden RTF untrusted-input limits, safe HTML output profiles, semantic-loss reporting, cancellation, structural editing, Word bridging, and performance baselines beyond the current tested profiles.
- [ ] Expand the provenance-recorded RTF producer corpus beyond current Word, Outlook, and LibreOffice evidence to Google Docs, macOS TextEdit/RTFD, EHR/CRM/helpdesk generators, and commercial libraries.
- [ ] Expand OpenDocument style, formula, drawing, embedded-content, and producer-corpus coverage while preserving unknown package content.

## Reader and document intelligence

- [ ] Improve `OfficeIMO.Reader.Pdf` reading-order, table reconstruction, asset/visual extraction, hierarchical chunking, confidence, and format-specific provenance against arbitrary-producer PDFs without creating a second PDF parser outside `OfficeIMO.Pdf`.

## Browser and agent surfaces

- [ ] Add reproducible browser-converter bundle-size, startup-time, and peak-memory gates for representative DOCX, XLSX, and PPTX conversions.
- [ ] Expose the shared conversion capability model consistently through package documentation, MCP discovery, and the browser converter.

## Completion rule

Remove an item when its public API, compatibility boundary, tests, generated evidence, and user documentation agree. GitHub Releases records delivered history, while `MIGRATION.md` retains only upgrade actions; this file does not retain completed milestones.
