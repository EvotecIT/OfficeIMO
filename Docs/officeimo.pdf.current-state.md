# OfficeIMO PDF current state

`OfficeIMO.Pdf` is the first-party, dependency-light PDF engine for creating, reading, inspecting, rendering, converting, and safely changing business PDFs. This document describes the current product contract. Open PDF work is tracked in the repository [roadmap](ROADMAP.md), and direct conversion evidence is generated in the [PDF conversion support matrix](officeimo.pdf-conversion-support-matrix.md).

## Ownership

- `OfficeIMO.Pdf` owns PDF parsing, writing, layout, rendering, security policy, signatures, forms, annotations, manipulation, diagnostics, and compliance evidence.
- `OfficeIMO.Core` owns reusable vector/raster primitives, text measurement and shaping seams, codecs, colors, paths, clipping, and image export policy; drawing APIs remain in the `OfficeIMO.Drawing` namespace.
- `OfficeIMO.Html` owns HTML/CSS parsing, resource policy, layout, pagination, and its backend-neutral scene. `OfficeIMO.Html.Pdf` maps that scene into PDF primitives.
- Word, Excel, PowerPoint, OpenDocument, Markdown, RTF, OneNote, AsciiDoc, LaTeX, Email, EPUB, and Visio packages remain source-format adapters over the shared PDF and Drawing owners.
- OCR execution remains an optional Reader/provider concern. PDF exposes image-only page evidence and accepts traced OCR results without embedding an OCR runtime.

The machine-readable [`pdf-conversion-scenarios.json`](pdf-conversion-scenarios.json) manifest owns direct-adapter routes, composition routes, fidelity status, and proof ownership. The checked-in conversion matrix is generated from that manifest and checked for drift.

## Public workflow

- `PdfDocument.Create(pdf => ...)` is the normal authoring entry point. `Compose(...)` appends through the same closed builder model; flow authoring is not duplicated on the root document. Incremental hosts can use `compose.Settings(...)` for the document-owned `PdfOptions` snapshot and `compose.Defaults(...)` for top-level page defaults without introducing a page boundary.
- `PdfDocument.Open(...)` is the normal read, inspect, and processing entry point.
- `PdfDocument.Read`, `Pages`, `Forms`, `Attachments`, `Bookmarks`, `Annotations`, `Stamp`, `Security`, `Redactions`, `Optimization`, and `Proof` expose focused workflow surfaces.
- `PdfDocument.Preflight(...)` provides non-throwing readiness and security evidence before a workflow is selected.
- `PdfDocument.Analyze(...)` provides the consolidated health, capability, diagnostic, optimization, signature, repair, mutation, and optional compliance report.
- `PdfDocument.CreateComplianceArtifact(...)` binds exact output bytes to internal readiness and external validator evidence.

Byte, stream, path, sync, and async entry points share the same engines. Caller-owned streams remain caller-owned, and asynchronous APIs are used for actual I/O rather than as wrappers around synchronous memory work.

## Workflow coverage

| Workflow | Current contract | Important boundary |
| --- | --- | --- |
| Create | Fluent flow and canvas APIs cover text, links, lists, tables, inline images/boxes, drawings, headers/footers, watermarks, metadata, sections, TOCs, replayable flow, columns, pagination controls, optional-content layers, portfolios, forms, tags, and viewer settings | Complex producer-specific layout is covered only when backed by a fixture and visible proof |
| Read and inspect | One bounded canonical parse exposes text, geometry, images, attachments, outlines, links, annotations, forms, actions, metadata, XMP, tags, layers, output intents, security, revisions, signatures, diagnostics, and compliance readback | Strict mode rejects defects; lenient mode reports explicit repairs and does not guess ambiguous intent |
| Merge, split, extract, reorder, rotate, resize, and copy pages | Shared import/rewrite engine with optional supported-annotation flattening and preservation reporting. Document-relative selectors cover absolute and reverse ranges, `last`, `odd`, `even`, and exclusions. | Catalog collisions, inherited resources, tags, forms, layers, and incremental structures remain subject to preflight |
| Crop | Any standard page boundary box can be changed | Setting `/CropBox` is not presented as destructive content removal |
| Stamp and watermark | Text, image, and imported PDF-page Form XObject stamps can target selected pages above or below content | Append-only stamping is available only when the signature and permission model allows it |
| Bookmarks and outlines | Generated nested outlines and destinations; existing outline read/preserve/split | Existing-document outline editing is limited and not a broad contract |
| Annotations | Create selected annotation types; inspect, filter, update a bounded subset, remove, flatten, and render authored or diagnosed synthesized appearances | Subtype geometry, threads, replies, attachment annotations, and redaction annotations remain explicitly bounded |
| Forms | AcroForm creation and existing-document field operations, appearances, calculation/tab order, typed/XFDF interchange, append-only values, and signature-field placement | XFA is detected and rejected rather than executed or silently changed |
| Password protection | Standard-security revisions 2–6, AES-256 default, AES-128 and explicit legacy RC4 modes, typed permissions, Unicode passwords, and owner-authorized rewrite workflows | Signed or security-sensitive mutation remains fail-closed when preservation proof is unavailable |
| Incremental updates | Shared writer for metadata/XMP synchronization, supported form values/appearances, annotation mutations, external-signature preparation, and DSS/VRI material; automatic mutation planning selects append-only output when the requested operation and preservation policy allow it | Encrypted incremental output and broader catalog operations are not claimed |
| Digital signatures | Approval, certification/DocMDP, and document-timestamp preparation; visible widgets; shared CMS/RFC 3161/X.509 processing through `OfficeIMO.Security`; DSS/VRI enrichment | Structure, signature math, digest, trust, revocation, time, permissions, and later revisions remain separate report states |
| Attachments and portfolios | Create, list, extract, add, replace, rename, remove, and preserve supported associated-file and collection structures | Existing portfolio presentation metadata editing remains bounded |
| Provenance | `PdfProvenance.Inspect`, `InspectFile`, `Remove`, and `RemoveFile` detect and selectively remove structurally valid C2PA Manifest Store associated files; cryptographic verification remains an optional `OfficeIMO.Security` provider concern | Defaults cap the source at 256 MiB, each manifest at 64 MiB, carriers at 128, structural/container entries at 65,536, and aggregate expanded data at 1 GiB; ambiguous associations, signatures, and unsupported rewrites fail closed |
| Optimization | Deterministic compression, unreachable-object removal, exact-stream/image/resource deduplication, object streams, Fast Web View, and action reporting | Optimization is a full rewrite and never claims signature preservation |
| Redaction | Reviewable geometry/search plans remove intersecting text, paths, annotations, fields, and supported image pixels; cleanup and residue checks are explicit | Unknown image encodings require a bounded caller decoder or fail-closed whole-placement policy |
| Render pages | Managed Drawing projection for supported paths, clipping, forms, images, shadings, patterns, appearances, alpha, blend modes, masks, fonts, and color spaces, including bounded ICC input transforms, destination-output soft proofing, rendering intents, and supported PDF color functions | Exact image, color, pattern, transparency, and font boundaries are maintained in the [image-export capability matrix](officeimo.image-export-capability-matrix.md#pdf-and-paged-source-adapters); unsupported cases remain explicit per-page diagnostics |
| Serialize | Buffered output with memory limits/spillover and opt-in forward-only object serialization to non-seekable destinations | Forward-only object writing is not forward-only layout; `ToBytes()` necessarily buffers the final artifact |
| Extract text and layout | Fast heuristic plus a pluggable understanding pipeline with stable JSON, Markdown, ALTO, hOCR, and PAGE XML | Editable reconstruction from arbitrary fixed-layout PDFs is not claimed |
| Compliance artifacts | PDF/A-2a/b/u, PDF/A-3a/b/u, PDF/A-4/4e/4f, PDF/UA-1, PDF/UA-2, Factur-X, and ZUGFeRD gates bind internal readiness and external evidence to exact bytes | Tags or metadata alone never establish conformance |

## Conversion and fidelity

Every adapter returns stable conversion evidence. `Faithful`, `FaithfulWithSubstitutions`, and `Degraded` are derived from structured warnings; producing a syntactically valid PDF alone is not a fidelity claim.

| Direction | Current contract |
| --- | --- |
| Office, OpenDocument, HTML, Markdown, RTF, OneNote, AsciiDoc, and LaTeX to PDF | Thin source adapters use the shared PDF/Drawing owners and the generated scenario manifest records the evidence level. Compatible Excel category-axis column/line/area combinations and scatter-axis series combinations flow into the shared renderer with secondary-axis assignment. Category-axis/scatter-axis mixtures are rejected with an explicit warning instead of silently omitting a series. |
| Normalized Reader result to PDF | Pages, blocks, tables, assets, links, forms, and diagnostics project through one explicit policy and merged evidence contract |
| PDF to Word | Metadata, page breaks, headings, paragraphs, lists, logical tables, links, supported images, and form placeholders are recovered when represented by the logical model. Stable warnings identify non-reconstructed outlines, tagged trees, optional-content groups, catalog/page actions, vectors, and non-link annotations. |
| PDF to Excel | Adjacent compatible tables can continue across pages with repeated multi-row headers suppressed; optional Boolean, date, percentage, and numeric typing is column-consistent. Bounded positioned-cell recovery is used only when tabular evidence is strong, and unrelated fixed-layout content remains explicit in the scope report rather than becoming editable cells. |
| PDF to PowerPoint | `PdfPowerPointImportOptions` defaults to `VisualPages`. `CreateEditableTables()` / `EditableTables` reconstructs detected tables, while `CreateHybrid()` retains each successfully rendered page as a visual layer and overlays editable tables at source-relative geometry. Reports distinguish visual-only content from omitted non-table page content, including failed hybrid page renders, and expose stable warnings for text, images, navigation, vectors, groups, forms/controls, annotations, and interactive media/animations. |
| PDF to HTML | Semantic and positioned-review profiles share the PDF logical/read model |
| PDF to PNG, JPEG, TIFF, WebP, or SVG | One page-to-Drawing projection serves authored and loaded documents, batches, and source-conversion results with budgets and diagnostics |

Word, Excel, and PowerPoint reference gates use pinned PDFs exported from the same checked-in fixtures with producer/version provenance, page geometry, source/reference hashes, semantic invariants, and recorded raster-distance budgets. A capability can be exact for a named invariant without claiming whole-document pixel equivalence.

OneNote conversion is a semantic-document projection, not a reconstruction of the free-form OneNote canvas. Email, EPUB, and Visio direct façades report their attachment/resource, pagination, preview, and semantic-fallback decisions rather than implying native-application equivalence.

## Deliberate boundaries and open product gaps

The engine already covers the common document-operation set expected from a
managed PDF library: authoring, page selection and composition, merge and
overlay, attachments, encryption, signatures, redaction, optimization and Fast
Web View, PDF 2.0, PDF/A and PDF/UA evidence, forms, annotations, structured
readback, and loss-aware format adapters. The remaining gap is mostly depth and
producer fidelity rather than a missing second API for the same operation.

The highest-value open areas are tracked in [ROADMAP.md](ROADMAP.md): difficult
remaining difficult Type 3 programs and color/pattern/transparency rendering, broader native Office layout,
externally proven standards and producer corpora, stronger hybrid editable
PDF-to-Office reconstruction, and optional provider packages for capabilities
that should not become core runtime dependencies. Browser JavaScript, TeX
execution, XFA execution, and automatic whole-document editable reconstruction
remain explicit non-contracts unless a separately owned, testable product shape
is adopted.

## Resources, fonts, and trust

The balanced default allows installed fonts plus bounded data URI and embedded-package resources. Arbitrary local-file reads and remote resolver calls remain disabled.

- `PdfResourcePolicy.CreatePortableDeterministic()` provides reproducible, host-independent behavior.
- `PdfResourcePolicy.CreateTrustedHost()` allows an application to opt into intended local or remote resolution.
- Conversion profiles control fidelity and content selection; they do not silently change trust.

Generated text can use registered TrueType or OpenType/CFF families without consuming the standard-font compatibility slots. Unavailable or non-embeddable fonts produce explicit substitution evidence. The optional `OfficeIMO.Drawing.HarfBuzz` adapter provides full OpenType shaping through the Drawing provider seam without adding HarfBuzz to the dependency-light core packages.

## Mutation and security policy

Preflight chooses a supported full rewrite, append-only revision, or refusal. Mutation reports prove byte-prefix, revision-chain, signature-range, DocMDP/FieldMDP, and preservation state where those properties matter.

- JavaScript and XFA are never executed.
- A signature object is not described as preserved when its signed byte range is invalidated.
- Catalog, page, resource, form, annotation, attachment, layer, tag, metadata, action, and security structures are not silently dropped during rewrite.
- Secure redaction, signature validity, and compliance require the appropriate cryptographic, residue, render, or external-validator evidence.
- Unsupported or ambiguous repairs remain diagnostics rather than guessed mutations.

## Performance and resource evidence

The PDF evidence lanes cover cold and cached analysis, SVG/PNG rendering, buffered serialization, forward-only object serialization, shaping, managed peak memory, largest transient buffers, deterministic bytes, and bounded hostile inputs. Limits include source bytes, objects, tokens, nesting, raw/decoded streams, content operations/operands, pages, page-tree depth, completed page/effect payloads, and serialized object buffering. Page rendering and image-export batches also honor one operation-wide deadline with typed timeout evidence while preserving caller-cancellation precedence.

Performance results apply to the measured input, options, runtime, and operating system. They do not substitute for correctness or preservation proof.

## Documentation rule

Package READMEs show supported public workflows. This file records the current cross-package PDF contract. The generated conversion matrix records route evidence, and [ROADMAP.md](ROADMAP.md) records open work. Dated comparison notes and implementation backlogs do not sit beside these sources of truth.
