# OfficeIMO PDF Current State And Roadmap

This is the canonical PDF product-state file. Keep it current. Do not add dated
PDF comparison or review snapshots beside it; fold durable conclusions into the
workflow inventory and implementation backlog below.

## Product Direction

`OfficeIMO.Pdf` should become the obvious dependency-light .NET library for
creating, reading, rendering, inspecting, converting, and safely changing
business PDFs.

The goal is not to copy another library's API or accumulate unrelated helpers.
The goal is to make common workflows easy while keeping difficult PDF behavior
explicit:

- normal files should have a short, fluent happy path;
- risky or unsupported files should fail closed with an actionable report;
- signed or append-sensitive files should use a proven incremental update path;
- full rewrites should preserve every supported document structure;
- visible output should have rendering proof, not only parser assertions;
- `OfficeIMO.Pdf` should not gain a runtime dependency on another PDF engine,
  browser, JavaScript runtime, or native renderer.

Word, Excel, PowerPoint, Markdown, HTML, RTF, and OneNote packages remain thin
adapters. Shared PDF parsing, writing, layout, rendering, security, signatures,
forms, annotations, resource trust, and manipulation belong in `OfficeIMO.Pdf`.
Reusable vector and raster primitives belong in `OfficeIMO.Drawing`. The
machine-readable direct-adapter and composition-route inventory is
[`pdf-conversion-scenarios.json`](pdf-conversion-scenarios.json); it is the
source of truth for supported routes and visual proof ownership.

Direct conversion uses one balanced resource default: installed fonts plus
bounded data URI and embedded-package resources are available for Unicode and
self-contained-document fidelity, while arbitrary local-file reads and remote
resolver calls remain disabled. Reproducible or untrusted pipelines can choose
`PdfResourcePolicy.CreatePortableDeterministic()`; applications that intentionally
resolve local or remote resources can choose `CreateTrustedHost()`. Profiles
control fidelity and content selection; they do not silently change trust. Zero-options and faithful Word or
Excel output also do not inject page numbers or worksheet-name headings that
were absent from the source.

Word conversion embeds distinct mapped document and run fonts while a PDF
family slot is available. The current writer has three such standard-family
slots; documents needing more receive a `NativeFontFamilySlotExhausted` loss
warning with the exact font and substitution instead of a silent alias.

OneNote conversion is deliberately named as a semantic-document projection.
It preserves hierarchy and reading-order content, reports canvas flattening and
asset placeholders, and does not claim to reproduce the free-form OneNote
canvas. PowerPoint has one stable native PDF shape renderer for full-slide
output; the shared visual snapshot remains the owner for PNG, SVG, HTML review,
and slide thumbnails, where its raster/vector scene contract is appropriate.

## Workflow Coverage

The status in this table describes the public workflow, not whether a low-level
PDF primitive exists somewhere in the codebase.

| Workflow | Status | Current contract | Work still needed |
| --- | --- | --- | --- |
| Create PDFs | Ready for common business documents | Fluent flow and canvas APIs cover text, links, lists, tables, images, drawings, headers, footers, watermarks, metadata, sections, generated TOCs, conditional/replayable flow, position capture, styled one-page containers, block-aware multi-column flow, generated optional-content layers, portfolios, form fields, tagging groundwork, and viewer settings. | Complex-script shaping, mixed inline boxes/images, line-level column balancing, widow/orphan and keep-with-next rules, multipage decorated containers, richer forms/annotations, and validator-backed compliance. |
| Read and inspect | Ready for common born-digital PDFs | Text, geometry, images, attachments, portfolio metadata, outlines, links, annotations, forms, actions, metadata, XMP, tagged content, layers, output intents, security, revisions, signature structure, and a bounded immutable raw-syntax projection are exposed. `PdfReadOptions.Limits` bounds input bytes, indirect objects, object characters/tokens/nesting, raw and decoded streams, content operations, page counts, and page-tree depth/nodes. Strict mode rejects structural defects; lenient mode records recovered versus detection-only findings for xref pointers, stream lengths, object boundaries/duplicates, page-tree counts/parents/kids, name trees, destinations, and unreachable semantic objects. | Continue adding producer-specific repair fixtures; never auto-repair a defect whose semantic intent is ambiguous. |
| Merge PDFs | Ready for rewrite-safe inputs | `PdfMerger` and `PdfDocument.MergeWith(...)` merge files, streams, or bytes; pages can be normalized and supported visual annotations flattened. | Explicit collision policies for forms, named destinations, page labels, outlines, attachments, metadata, and catalog state; broader complex-file proof. |
| Split and extract pages | Ready for rewrite-safe inputs | Single pages, page ranges, range expressions, fixed-size groups, and bookmark-derived ranges are supported. | Better preservation policy reporting for structures whose targets fall outside the selected pages. |
| Remove, duplicate, move, reorder, and rotate pages | Ready for rewrite-safe inputs | Fluent and static APIs cover the standard page-editing operations. `ComposePages`/`ComposePageRanges` allow selected subsets and repetitions through the shared extraction engine; convenience APIs reverse documents, repeat selections, and round-robin interleave even or uneven ranges. | Broader object-stream, tagged, layered, form-heavy, attachment-heavy, and incremental-file proof. |
| Copy pages from another PDF | Ready for rewrite-safe inputs | Pages can be appended, prepended, or inserted from another PDF, with optional annotation flattening. | The same collision and catalog policies needed by merge, plus a concise import report. |
| Resize pages | Ready for the supported rewrite subset | Pages can be resized with fit/fill/stretch behavior and destination transforms. | Broader preservation and visual corpus for inherited resources and unusual page trees. |
| Crop pages | Partial | Any production boundary box can be set, including `/CropBox`, `/TrimBox`, `/BleedBox`, `/ArtBox`, and `/MediaBox`. | Add named crop APIs, crop-and-translate, and an explicitly destructive crop mode that removes or clips content outside the retained area. Setting `/CropBox` alone must not be described as content removal. |
| Stamp and watermark | Ready for the supported rewrite subset | Text and image stamps/watermarks can target selected pages through fluent or static APIs. Complete source PDF pages can be imported as vector-preserving Form XObjects above or below selected target content with fit, alignment, rectangle, and opacity controls. | Rich text, reusable appearance templates, optional append-only stamping when signature permissions allow it, and wider resource-preservation proof. |
| Bookmarks and outlines | Partial | Generated documents can create nested outlines and named destinations; existing outlines can be read, preserved when supported, and used to split a document. | Add an existing-document outline editor: add, remove, rename, move, nest, retarget, rebuild from headings, and validate broken destinations. |
| Annotations | Partial | Generated PDFs can create text, free-text, highlight, and link annotations. Existing annotations can be read, filtered, flattened for supported appearances, updated in a small metadata/style subset, or removed. Updates and removals can use append-only revisions for unsigned/approval-signed inputs and certification signatures with DocMDP `/P 3`; `/P 1` and `/P 2` are blocked, and widget edits remain routed through the FieldMDP-aware form engine. Results expose the mutation plan plus rewrite-preservation or signature/revision proof. | Add annotations to existing pages; move/resize them; edit subtype-specific geometry and appearance; reply/thread support; selective flattening; broader markup, ink, stamp, file-attachment, and redaction annotation behavior. |
| Password protection | Ready for supported Standard-security workflows | Generated and rewritten PDFs default to AES-256 revision 6, with AES-128 interoperability and explicit legacy RC4 modes, typed permissions, Unicode password handling, revision 2-6 reading, authenticated user/owner roles, and owner-authorized encrypt/decrypt/re-encrypt workflows with preservation reports. | Expand encrypted mutation coverage beyond the dedicated security rewrite and keep signed/security-sensitive inputs fail-closed. Certificate signing and validation stay in the optional first-party cryptography package. |
| Metadata | Partial | Info-dictionary title, author, subject, and keywords can be replaced or updated by full rewrite or append-only revision. Generated PDFs can emit XMP and profile metadata; existing XMP is readable. | Edit and synchronize Info plus XMP, preserve custom schemas, manage dates/producer/creator deliberately, and report conflicts instead of silently choosing one source. |
| Forms | Broad | Generated and existing-document AcroForms support field creation, rename, remove, move, defaults, flags, calculation and tab order, appearance regeneration, exact-field flattening, typed/XFDF data interchange, append-only value updates, and empty signature-field placement. XFA is detected and explicitly rejected by the AcroForm editor rather than executed or silently changed. | Expand field kinds and appearance fidelity only when backed by interoperable fixtures; keep XFA outside the dependency-light core. |
| Incremental updates | Partial but real | A shared incremental object writer appends metadata, supported form values/appearances, external-signature preparation, and DSS/VRI validation material without replacing prior bytes. It preserves object generations and trailer state and emits classic or xref-stream revisions. Mutation plans and before/after reports prove byte-prefix, revision-chain, signature-range, and DocMDP/FieldMDP state. | Add encrypted incremental serialization, supported annotation/stamp/catalog operations, and broader interoperability fixtures. |
| Digital signatures | Partial | Approval, certification/DocMDP, and document-timestamp profiles can prepare external signatures; approval/certification fields can have visible widget appearances. A signer callback accepts CMS/CAdES/RFC 3161 bytes from cloud, HSM, smart-card, or local implementations without moving key storage into the PDF core. The optional first-party `OfficeIMO.Pdf.Cryptography.Pkcs` package has no outside runtime package and provides RSA/SHA-256 detached CMS signing plus managed CMS/RFC 3161 parsing, signature/digest validation, X.509 chains, caller trust callbacks, and revocation policy on every supported target. After signature math and digest verification, `PdfLongTermValidationEnricher` can append DER certificate, OCSP, and CRL streams in an ETSI DSS/VRI revision while retaining all earlier bytes and evidence. Reports keep structure, math, digest, trust, revocation, time, permissions, and later revisions separate. | Add non-RSA signing algorithms, deeper timestamp/revocation fixtures, managed rendering proof for signature widgets, and external interoperability proof for B-LT/B-LTA workflows without claiming conformance prematurely. |
| Attachments and portfolios | Broad | Generated associated/embedded files and collection dictionaries are supported, including portfolio fields, sort order, initial document, and view. Existing attachments can be listed, selectively extracted, added, replaced, renamed, removed, and edited through the attachment engine; portfolio metadata is readable and supported rewrites retain it when preservation proof passes. | Add a focused existing-portfolio metadata editor and broaden viewer/interoperability fixtures for collection presentation. |
| Optimization | Broad | Deterministic Balanced, MaximumCompression, Web, Archival, and Custom profiles support lossless stream compression, unreachable-object removal, exact-stream and decoded-image deduplication, font/resource dictionary deduplication, classic or xref-stream output, object-stream packing, keep-original-if-larger behavior, per-action reporting, and post-save preservation proof. The Web profile emits standards-compliant Fast Web View output with two cross-reference sections plus page-offset and shared-object hint tables; linearization deliberately requires classic cross-reference tables without object streams. | Expand semantic deduplication and linearization corpus coverage only with bounded decoders and interoperable fixtures; optimization remains an explicit full rewrite and never claims signature preservation. |
| Redaction | Secure workflow available | Reviewable geometry/search plans remove intersecting text, vector paths, annotations, form fields, and image pixels. Built-in image normalization covers transformed placements, indexed/color-key/explicit/soft masks, and clone-on-write reuse; JPEG and other codecs use an optional bounded decoder contract or an explicit fail-closed/whole-placement policy. Cleanup policies cover metadata, attachments, structure/alternate text, and optional content. Proof combines extraction, raw/decoded residue checks, managed rendering, and pluggable independent validators. | Expand the hostile/corpus fixture set as new producer-specific encodings are found. |
| Render PDF pages | Broad managed subset with explicit diagnostics | Static pages project to shared Drawing primitives with paths/clipping, forms, images, axial/radial shadings, colored and basic uncolored vector tiling fills, supported annotation/form appearances, alpha, standard blend modes, Form-XObject alpha/luminosity soft masks, exact embedded TrueType outlines, and managed CMYK/Lab plus simplified calibrated-gray/RGB conversion through `OfficeIMO.Drawing`. The shared Drawing raster and SVG paths own reusable tiling, blending, and masking. PNG/SVG batches provide ranges, DPI/scale/background, thumbnails, cancellation, budgets, and per-page reports. A generated manifest reports simplified/unsupported operators and resources, and optional image codecs plug into shared Drawing rasterization without becoming core dependencies. | Extend fidelity from corpus failures while keeping CFF/Type 3 gaps, resource-specific calibrated parameters, unsupported ICC spaces, stroked/text tiling patterns and other broader pattern edge cases, and incomplete layer projection explicit in page reports. |
| Serialize generated PDFs | Bounded payload streaming | `PdfOptions.PageContentMemoryLimitBytes` bounds completed page/effect content retained during layout, and `PdfOptions.ObjectBufferMemoryLimitBytes` bounds completed indirect-object bytes during serialization. Both stores spill excess payloads to indexed temporary files; large stream objects are spooled without a duplicate combined buffer, and final stream assembly copies spilled objects in bounded chunks for plain and encrypted saves. Spill files are removed on disposal. | Per-page metadata and the authored block model remain proportional to document size, the active page is materialized while it is processed, and `ToBytes()` necessarily buffers the final artifact. Fully forward-only layout/output needs a deeper writer contract and representative memory gates. |
| Text and layout extraction | Broad, strategy-driven | The fast heuristic remains the default. A pluggable six-stage understanding pipeline provides confidence/evidence and stable JSON, Markdown, ALTO, hOCR, and PAGE XML. The built-in advanced profile adds rotation/arbitrary-baseline grouping, spatial and non-rectangular regions, multi-column/spanning-band order, tables, captions, headers/footers, and footnotes. | Refine advanced heuristics from real mixed-layout corpora and use provider stages for domain-specific reconstruction rather than hard-coding every document family. |
| PDF to Office/HTML/data | Partial by design | PDF-to-HTML review output, table export, Reader chunks, and limited PowerPoint table import use the shared logical model. | Improve the logical model and confidence/proof first. Do not promise general editable reconstruction from a presentation format. |
| Office/HTML/Markdown/RTF to PDF | Broad but evolving | Thin adapters use the shared PDF and Drawing engines. HTML uses the shared render scene introduced by the HTML/PDF/image work. | Continue converter-specific fidelity only when the missing primitive is truly source-specific; otherwise improve the shared PDF, Drawing, or HTML owner. |
| PDF/A, PDF/UA, and e-invoices | Groundwork only | Output intents, tagging, XMP identification, associated files, Factur-X/ZUGFeRD groundwork, and compliance proof reports exist. | Pass an external validator for one narrow profile before making a conformance claim, then expand profile by profile. |

## Current Architecture To Keep

- `PdfDocument.Create(...)` is the normal document-authoring entry point.
- `PdfDocument.Open(...)` is the normal fluent read and processing entry point.
- Static manipulation types remain useful for direct byte, stream, and path
  workflows, but should feed the same engines and reports as the fluent API.
- `PdfReadDocument` and the logical models are the parser/read source of truth.
- `OfficeIMO.Drawing` is the shared managed scene, SVG, and raster owner.
- `OfficeIMO.Html` owns HTML/CSS parsing, resource policy, layout, pagination,
  and its backend-neutral render scene.
- `OfficeIMO.Html.Pdf` only maps that render scene to PDF primitives.
- OCR execution stays outside `OfficeIMO.Pdf`; the core exposes image-only page
  evidence and accepts traced OCR results through Reader workflows.

## Implementation Backlog

The order matters. Convenience APIs should not get ahead of preservation and
security proof.

### P0 - Make Mutation Safety A Product Contract

- [x] Add a single mutation planner used by every editing API. It should return
  `FullRewrite`, `AppendOnly`, or `Blocked`, list the exact structures and
  permissions that drove the decision, and name the proof required after the
  operation.
- [x] Replace operation-specific safety guesses with shared capability records
  for page-tree changes, content changes, catalog changes, form changes,
  annotations, metadata, attachments, encryption, and signatures.
- [x] Build a curated interoperability corpus containing classic xref tables,
  xref streams, object streams, hybrid references, incremental revisions,
  unusual generations, linearized files, signed/certified files, encrypted
  files, tagged PDFs, optional-content layers, attachments, name trees,
  output intents, active content, and complex forms.
- [x] Run every rewrite scenario through `PdfRewritePreservationMatrix` and
  retain the original blocker or preservation report. Expected blockers must be
  distinguishable from regressions.
- [x] Add external development-time checks for syntax, rendering, signatures,
  and the selected compliance profile. External tools remain test/build proof,
  not runtime dependencies.
- [x] Fuzz the tokenizer, object parser, stream decoders, page-tree traversal,
  content parser, form/annotation parsers, and incremental revision reader with
  strict input, recursion, object-count, decoded-byte, and time budgets.

Exit criterion: every public mutation can explain why it will rewrite, append,
or refuse the input, and the decision is exercised against the curated corpus.

### P0 - Generalize Append-Only Updates

- [x] Extract the current metadata, form, and signature revision logic into one
  append-only writer that can add and replace indirect objects, preserve object
  generations, emit classic or xref-stream revisions, maintain `/Prev`, trailer
  references, file identifiers, and encryption context, and leave all prior
  bytes unchanged.
- [x] Model DocMDP and FieldMDP permissions per requested operation and target,
  not only as document-wide flags.
- [x] Add before/after signature reports that identify which signatures still
  cover which revisions and whether the requested change is permitted.
- [x] Add operations incrementally: metadata/XMP first, then supported form
  updates, annotation updates, permitted stamps, DSS/LTV material, and finally
  other catalog changes that have fixture-backed proof.
- [x] Never route page deletion, page import, destructive redaction, encryption
  changes, or other structurally incompatible operations through append-only
  mode merely to keep an old signature object present.

Exit criterion: append-only output is byte-prefix identical to the input, has a
valid revision chain, and retains the expected signature/permission state.

### P0 - Modern Security And Signature Validation

- [x] Add typed standard-security algorithms and permissions. Make AES-256
  revision 6 the normal modern output, support AES-128 where interoperability
  requires it, and keep RC4 only behind an explicit legacy option.
- [x] Support reading revision 5/6 password security with correct Unicode
  password processing and metadata-encryption behavior.
- [x] Add encrypt, change-password/permissions, and owner-authorized decrypt
  workflows for existing PDFs, with full preservation reports.
- [x] Put CMS signature verification behind a small cryptography seam so the
  dependency-free PDF parser owns byte ranges and signed attributes while the
  optional first-party cryptography package owns bounded DER/CMS parsing,
  signature math, certificate-chain, timestamp, OCSP, and CRL policy without an
  external runtime package.
- [x] Add signing profiles for approval signatures, certification signatures,
  document timestamps, visible signature appearances, and external/cloud/HSM
  signers without placing key-storage logic in `OfficeIMO.Pdf`.
- [x] Add DSS/VRI creation and append-only LTV enrichment only after
  cryptographic verification and revision proof are in place.

Exit criterion: a caller can create a modern encrypted PDF, safely decrypt or
re-encrypt an existing supported PDF, and obtain a cryptographic signature
report that clearly separates mathematical validity, trust, revocation, time,
permissions, and later revisions.

### P1 - Finish The Standard Editing Workflows

- [x] Add one merge/import policy model for document metadata, outlines, named
  destinations, page labels, AcroForm field-name collisions, annotation
  destinations, attachments, output intents, layers, viewer settings, and page
  size normalization. Every non-trivial choice must be reported.
- [x] Add convenience workflows for interleaving, repeating, reversing, and
  composing selected page ranges without bypassing the shared import engine.
- [x] Add `SetCropBox(...)`, `SetTrimBox(...)`, and related named page-box APIs,
  followed by crop-and-translate and explicitly destructive content cropping.
- [x] Add a bookmark editor for add/remove/rename/move/nest/retarget/rebuild,
  plus broken-target validation.
- [x] Add destination-conflict handling for bookmarks during page edits and
  merges.
- [x] Expand annotation editing to create annotations on existing pages,
  update rectangles/quads/vertices/ink paths/line endings/popups/replies,
  regenerate appearances, remove actions, flatten selected annotations, and use
  append-only changes when the mutation planner permits them.
- [x] Expand metadata editing to Info and XMP with synchronized common fields,
  custom-schema preservation, explicit clear/preserve semantics, and both full
  rewrite and append-only variants.
- [x] Add an existing-document AcroForm editor for field creation, rename,
  remove, move, default values, flags, calculation order, tab order, selected
  flattening, data import/export, and signature-field placement.
- [x] Add attachment editing for embedded files and associated files, including
  relationship, description, MIME type, dates, checksum, replace, rename,
  remove, and safe file-name extraction.
- [x] Add a sanitization workflow that can remove or quarantine JavaScript,
  launch actions, remote navigation, submit/import actions, embedded files,
  rich media, and unsafe URI schemes, then prove the active-content inventory
  is empty or matches an allow list.
- [x] Extend lossless optimization with image/resource/font deduplication,
  object/xref stream output, deterministic profiles, and standards-compliant
  Fast Web View linearization with page and shared-object hint tables.
  Optimization remains a declared full rewrite and does not imply signature
  preservation.

Exit criterion: the common paid-library workflows are available through one
readable `PdfDocument.Open(...)` pipeline and return a preservation/security
report instead of silently discarding document structures.

### P1 - Complete Redaction As A Security Feature

- [x] Finish partial-image rewriting for JPEG, transformed placements, color-key
  masks, explicit masks, soft masks, indexed colors, and shared/reused resources.
- [x] Remove or rewrite intersecting text, paths, annotations, structure-tree
  references, alternate text, optional-content references, metadata, and
  attachments according to an explicit redaction policy.
- [x] Add search-driven redaction for literal text, regular expressions, logical
  fields, and caller-provided geometry, with reviewable plans before apply.
- [x] Keep visual-only overlays as a separately named non-redaction operation.
- [x] Validate output through extraction, raw/decoded-stream residue checks,
  page rendering, and at least one independent development-time parser.

Exit criterion: the redaction API can state and prove what was removed; no
visual-only operation is allowed to claim secure redaction.

### P2 - Make Rendering And Viewing Broadly Useful

- [x] Turn the current managed page renderer's supported subset into a generated
  capability manifest with stable diagnostics for every skipped or simplified
  operator/resource.
- [x] Establish corpus-driven static rendering coverage for stream filters and image
  codecs, Type 1/3/TrueType/OpenType/CID fonts, color spaces and ICC handling,
  tiling and shading patterns, functions, transparency groups, masks, blend
  modes, clipping, form XObjects, annotations, AcroForm appearances, and layers,
  with either a tested projection, an optional provider seam, or a stable
  per-page unsupported/simplified diagnostic for every category.
- [x] Add page-range rendering, DPI/scale/background options, PNG/SVG batches,
  thumbnails, cancellation, render limits, and per-page reports.
- [x] Add text-selection and hit-testing primitives over glyph geometry,
  annotation/link/form hit regions, page transforms, and selection quads.
- [x] Keep WPF, WinUI, MAUI, Avalonia, Blazor, and other viewer controls outside
  the core. They should be thin packages over shared rendering, selection,
  navigation, and caching contracts if real consumers require them.
- [x] Keep platform printing outside the core as well. A thin print adapter must
  honor document permissions and reuse the same page geometry and render plan;
  it must not introduce another parser or silently rasterize at low quality.
- [x] Add visual comparison helpers for page alignment, pixel/structural diffs,
  ignored regions, thresholds, and a human-review gallery.

Exit criterion: arbitrary static business PDFs either render correctly or
return a precise unsupported-feature report; UI packages do not contain a
second parser or renderer.

### P2 - Deepen Document Understanding

- [x] Split text understanding into pluggable stages: glyph decoding, word
  grouping, line grouping, page segmentation, reading order, and semantic
  classification.
- [x] Add strategies for rotated text, arbitrary baselines, multiple columns,
  L-shaped regions, tables, captions, headers/footers, footnotes, and mixed
  drawing/text layouts. Keep the current lightweight heuristic as the fast
  default.
- [x] Add confidence and diagnostic evidence to words, lines, regions, tables,
  headings, lists, and inferred reading order.
- [x] Export the same structured model to stable JSON, Markdown, ALTO XML, hOCR,
  and PAGE XML without duplicating extraction algorithms in each exporter.
- [x] Add a debug overlay that renders word/line/region boxes and reading order
  through the shared Drawing/PDF annotation primitives.
- [x] Keep OCR as a provider interface and merge contract. Do not ship an OCR
  engine or model in the dependency-light PDF core.

Exit criterion: callers can choose a fast or advanced layout strategy, inspect
why content was ordered or classified, and export a standard interchange model.

### P2 - Parser And Repair Diagnostics

- [x] Add a strict/lenient parsing policy with a repair report. Lenient mode may
  recover known structural defects but must never silently change semantic or
  security behavior.
- [x] Diagnose and, where safe, rebuild broken xref tables/streams, malformed
  page trees, incorrect stream lengths, orphaned objects, duplicate object
  identifiers, invalid name trees, and broken destinations.
- [x] Add decoded-stream and object-count budgets before allocating large
  buffers; report compressed and decoded sizes separately.
- [x] Add a PDF debugger dump for objects, revisions, page resources, content
  operators, reachability, and optionally decoded streams or a decompressed
  inspection copy. Keep it a diagnostic projection, not another mutable object
  model.

Exit criterion: malformed inputs produce bounded, actionable diagnostics and
any repair is explicit, reproducible, and validated before save.

### P3 - Deepen Creation And Conversion

- [x] Complete dependency-free OpenType/CFF charstring subsetting for generated
  and supported appearance-font output.
- [ ] Supply or build an optional implementation of the existing
  `IPdfTextShapingProvider` seam for full bidirectional layout, contextual
  shaping, mark positioning, and script-specific substitution. Keep it outside
  the dependency-free core; the built-in engine continues to provide Unicode
  scalar and Latin-ligature modes with explicit diagnostics.
- [x] Add shared block-aware multi-column flow, styled one-page containers,
  conditional/replayable flow constraints, position capture, semantic sections,
  generated TOCs, and optional-content layers before adapter-specific variants.
- [ ] Continue with line-level column balancing, mixed inline boxes/images,
  keep-with-next, widow/orphan, absolute-layout, multipage decorated containers,
  drawing, and deeper pagination behavior.
- [ ] Keep HTML/CSS fidelity work in the canonical HTML/PDF/image plan. The PDF
  package should only add missing PDF-native primitives or writer support.
- [ ] Add reusable typed report components and recipes only after repeated real
  examples identify a stable abstraction. Do not introduce a second template
  language or dependency-injection requirement into the PDF core speculatively.
- [ ] If real invoice, label, ticket, or logistics workflows require barcodes or
  QR codes, implement the encoding and drawing once in the smallest reusable
  `OfficeIMO.Drawing` owner (or a narrow optional adapter). PDF, Word, Excel,
  PowerPoint, and HTML surfaces should place the same typed drawing primitive.
- [ ] Expand generated annotations, form fields, signature appearances,
  attachments, layers, navigation, tagged structure, and accessible names
  through the same primitives used by existing-document editing.
- [ ] Continue Word, Excel, PowerPoint, Markdown, HTML, RTF, and OneNote fidelity
  through shared primitives first, with conversion warnings and proof snapshots
  preserved through post-processing.
- [ ] Promote AsciiDoc, LaTeX, and OpenDocument manual compositions to direct
  adapters only when their existing projection diagnostics flow automatically
  into the final PDF result. Do not advertise email, EPUB, or Visio as direct
  PDF conversion until body/asset/book/vector-page policies are explicit and
  visually proven.

Exit criterion: new fidelity work improves the shared engine or has a documented
source-format reason to remain in a thin adapter.

### P3 - Earn Narrow Compliance Claims

- [ ] Choose one small PDF/A target, wire its external validator into proof, and
  fix every requirement without hiding warnings.
- [ ] Add PDF/UA only when structure ownership, reading order, alternate text,
  annotations, forms, links, language, fonts, and Unicode mapping can be proven.
- [ ] Validate Factur-X/ZUGFeRD attachment, XMP, output-intent, and profile rules
  as a complete artifact rather than independent dictionary checks.
- [ ] Store validator version, profile, result, and artifact hash in the proof
  report so CI evidence is reproducible.

Exit criterion: a profile is advertised only when internal readiness and the
external validator both pass the exact generated artifact.

### P4 - Productize The Library Surface

- [ ] Make byte, stream, path, sync, async, and fluent overloads consistent for
  every supported operation without multiplying independent implementations.
- [ ] Preserve input/output diagnostics through chained operations and expose a
  final pipeline report with mutation decisions, warnings, preservation proof,
  hashes, page counts, and timings.
- [x] Bound completed page/effect content and serialized-object retention for
  stream saves with independent memory thresholds, indexed temporary-file
  spillover, direct large-stream spooling, and chunked final assembly, including
  encrypted output.
- [ ] Add fully forward-only layout and serialization before describing async
  APIs as fully streaming. Per-page metadata and the authored block model remain
  proportional to document size, while `ToBytes()` buffers the final artifact.
- [ ] Generate the public support matrix and README examples from tested
  capability records so documentation cannot drift from the implementation.
- [ ] Keep NativeAOT, trimming, deterministic-output, cross-platform, memory,
  and performance gates for representative small, large, hostile, and
  incrementally updated PDFs.
- [ ] Isolate optional dependencies in narrow adapter packages. The core should
  expose provider contracts for cryptography, OCR, advanced shaping, or codecs
  only where the BCL and current OfficeIMO projects cannot supply the behavior.

Exit criterion: normal workflows are discoverable from `PdfDocument`, advanced
behavior remains composable, and package dependencies reflect real optional
capabilities rather than leaking into every consumer.

## Guardrails

- Do not add another general PDF library as the implementation behind
  `OfficeIMO.Pdf`.
- Do not preserve a signature object while invalidating its byte range and call
  that signature preservation.
- Do not silently drop catalog, page, resource, form, annotation, attachment,
  layer, tagged-content, metadata, action, or security structures during a
  rewrite.
- Do not execute JavaScript or XFA. Inspect, report, remove, preserve when safe,
  or route elsewhere.
- Do not put OCR models, browser processes, UI frameworks, or native renderer
  dependencies in the core package.
- Do not create separate layout, rendering, or proof engines for each converter
  or UI surface.
- Do not claim secure redaction, PDF/A, PDF/UA, signature validity, or fidelity
  from dictionary presence or unit tests alone.
- Do not turn every discovered idea into an abstraction. Add the smallest
  reusable owner after a real workflow and proof path exist.

## Definition Of A Universal OfficeIMO PDF Workflow

A workflow is complete when:

1. the public API is short and discoverable for the normal case;
2. the same engine serves byte, stream, path, fluent, sync, and async surfaces;
3. preflight chooses full rewrite, append-only update, or refusal explicitly;
4. unsupported content and preservation risks have stable diagnostics;
5. output is re-read and compared against the operation's preservation policy;
6. visible changes have managed render proof and independent proof where risk
   warrants it;
7. security, signature, redaction, and compliance claims use the required
   cryptographic or external validation; and
8. the capability remains in the dependency-light shared owner instead of being
   copied into a converter, UI, wrapper, or example.

## Documentation Rule

Update this file when current behavior or priorities change. Package READMEs
should show supported public workflows; generated command/API documentation
should come from code metadata. Remove obsolete investigation notes after their
durable conclusions are represented here.
