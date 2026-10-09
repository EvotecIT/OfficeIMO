# DjVu reading, rendering and conversion design

This design defines the planned managed DjVu document contract. The [product roadmap](ROADMAP.md#djvu-scanned-documents) owns its open deliverables. No DjVu reader or decoder is implemented by this design.

Goal: read `.djvu` and `.djv` books, extract existing text layers, render selected pages and convert them to PDF, with explicit OCR only where needed. Implement managed codecs in OfficeIMO without a new external runtime dependency. DjVuLibre may serve as an isolated validation oracle; neither an installed executable nor a vendored third-party decoder becomes the product implementation. DjVu authoring is outside this read/render/convert goal.

### Owner and planned API

`OfficeIMO.DjVu` owns the bounded IFF/DjVu container, component identity, page model, BZZ/ZP text decompression, JB2 symbol decoding, IW44 colour decoding and layer composition. Reuse `OfficeIMO.Core` raster surfaces, image encoders and resource helpers. Keep PDF output in `OfficeIMO.DjVu.Pdf` over `OfficeIMO.Pdf`, and semantic ingestion in `OfficeIMO.Reader.DjVu` over Reader.Core. Reuse Reader.Ocr and the existing `IOcrEngine` contract for enrichment; hosts must not grow their own parsers, decoder processes or OCR policies.

The planned public entrypoints follow existing document owners:

| Surface | Contract |
| --- | --- |
| `DjVuDocument.Load(byte[]/Stream/string, DjVuReadOptions?, CancellationToken)` | One immutable owned source snapshot; stream input starts at its current position, supports non-seekable streams and stays open; file input retains no handles. Options are snapshotted. |
| `DjVuDocument.Pages` | Read-only pages in source document order, with stable component identity, one-based page numbers, pixel dimensions, DPI and rotation. Shared dictionaries have one owner per source document. |
| `DjVuPage.GetText(CancellationToken)` returning `DjVuTextResult` | Stored text, validated text-zone hierarchy and source offsets, normalized page bounds, and diagnostics. Distinguish absent, empty and corrupt layers; do not describe stored OCR text as newly recognized or authoritative. |
| `DjVuPage.Render(DjVuRenderOptions?, CancellationToken)` returning `DjVuRenderResult` | An `OfficeRasterImage` plus diagnostics; explicit resolution, selected region, orientation, background and hard pixel/allocation limits. Report unsupported layers instead of returning an apparently complete blank page. |
| PDF conversion extensions | Byte/stream output and report-returning operations over OfficeIMO.Pdf; preserve source page order and physical size, compose decoded images and qualified searchable text, and report text/geometry loss. File helpers delegate to the same operation. |
| `AddDjVuHandler(...)` | Thin Reader adapter exposing text, zones and requested raster assets, with capability qualifications for the implemented codec profiles. Reader.Ocr retains existing source/recognition provenance and aggregate budgets. |

Default file handling is self-contained. Bundled components are resolved by exact source identity; indirect documents and `INCL` references may not trigger filesystem or network reads. Add indirect input only through a caller-supplied, byte-returning resolver with explicit count/byte/depth budgets, no implicit directory access, and traversal/cycle/duplicate-identity proof. Unsupported secure/encrypted variants fail explicitly. `DjVuReadOptions` must bound source bytes, pages, chunks/nesting, expanded bytes, text/zones, symbol dictionaries and decoded pixels before allocation, with aggregate limits and cancellation inside decompression and rendering loops.

### Acceptance contracts

These contracts define evidence for the roadmap deliverables; they are not a separate backlog.

| Deliverable | Required evidence |
| --- | --- |
| DJ0 — Format and codec feasibility | freeze redistributable, hash-bound independent single-page and bundled fixtures for `TXTa`/`TXTz`, shared `DJBZ` dictionaries, `Sjbz`, progressive `BG44`, foreground palettes and rotation. Include compressed directory/navigation data, malformed lengths and hostile expansion cases. Establish expected text, zones, page order and reference rasters, plus the codec/specification provenance and licensing boundary before implementation. An oracle-generated round trip alone does not qualify archival inputs. |
| DJ1 — Useful text reader | implement container/page discovery and bounded BZZ/ZP decoding for existing compressed and uncompressed text, with zone offset/geometry validation, component identity and explicit missing/corrupt-layer behavior. Deliver the byte/stream/path API and Reader text adapter. Acceptance requires exact expected text and order from independent books, multibyte UTF-8 offset tests, caller-stream ownership, cancellation and enforced expansion limits on every codec path. Metadata-only parsing or `TXTa`-only extraction does not close this milestone. |
| DJ2 — Scanned-page rendering | implement bounded JB2/shared-dictionary and IW44/progressive decoding, palettes, layer composition, DPI and rotation on the existing raster surface. Acceptance requires independently rendered all-page comparisons for monochrome, colour and compound scans, non-white content/geometry coverage, selected-page/region/resolution behavior, safe malformed-input rejection and cancellation/resource-release evidence. A substitute image or unsupported blank page does not close the milestone. |
| DJ3 — PDF and optional OCR | use the qualified decoder and OfficeIMO.Pdf writer for raster pages and searchable existing text. Reuse the OCR pipeline for an explicit missing-text-only policy; preserve provenance/confidence and do not overwrite existing text without a requested replacement policy. Acceptance requires independent PDF reopen/render and text-search checks, physical page size/order, text-to-image geometry, mixed text/scanned pages, diagnostics for unavailable OCR, cancellation and bounded memory. Successful PDF serialization alone does not establish content preservation. |
| DJ4 — Consumer and package qualification | expose only implemented profiles through the Reader/conversion catalogs and thin Workflows/Tool/Studio consumers. Qualify source and packed consumers on supported .NET/platform targets, representative native/browser paths where actually supported, large-book opt-in memory/performance runs and runtime-dependency inventories. Move current contracts to package READMEs/support matrices and remove completed work here after independent review and PR validation. |

Text extraction can be qualified independently. Rendering, PDF and OCR capabilities require their respective codec and artifact evidence. A container parser, uncompressed-text-only reader or external-tool adapter does not satisfy the complete read/render/convert contract.
