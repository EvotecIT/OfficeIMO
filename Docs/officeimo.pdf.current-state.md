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

Word, Excel, PowerPoint, OpenDocument, Markdown, HTML, RTF, OneNote, AsciiDoc, and LaTeX
packages remain thin adapters. AsciiDoc and LaTeX reuse their existing
loss-aware Markdown projections and the Markdown PDF renderer; they do not add
format-specific layout engines. Shared PDF parsing, writing, layout, rendering,
security, signatures, forms, annotations, resource trust, and manipulation belong in `OfficeIMO.Pdf`.
Reusable vector and raster primitives belong in `OfficeIMO.Drawing`. The
machine-readable direct-adapter and composition-route inventory is
[`pdf-conversion-scenarios.json`](pdf-conversion-scenarios.json); it is the
source of truth for supported routes and visual proof ownership. The checked-in
[`PDF conversion support matrix`](officeimo.pdf-conversion-support-matrix.md) is
generated from that manifest and verified for drift in CI.

OpenDocument text, spreadsheet, and presentation callers use one direct
loss-aware façade over the existing semantic and PDF adapters. It combines
projection-stage and PDF-stage diagnostics without adding another layout or
rendering engine.

Direct conversion uses one balanced resource default: installed fonts plus
bounded data URI and embedded-package resources are available for Unicode and
self-contained-document fidelity, while arbitrary local-file reads and remote
resolver calls remain disabled. Reproducible or untrusted pipelines can choose
`PdfResourcePolicy.CreatePortableDeterministic()`; applications that intentionally
resolve local or remote resources can choose `CreateTrustedHost()`. Profiles
control fidelity and content selection; they do not silently change trust. Zero-options and faithful Word or
Excel output also do not inject page numbers or worksheet-name headings that
were absent from the source.

Generated PDF text can use any number of registered named TrueType or
OpenType/CFF families without consuming the three standard-font compatibility
slots. Word, Excel, PowerPoint, HTML, headings, tables, lists, headers, and
footers preserve named-family selection when an embeddable font is available.
Unavailable or non-embeddable source fonts fall back to a mapped PDF font and
retain an explicit conversion warning rather than a silent alias.

Word, Excel, and PowerPoint fidelity is measured against pinned PDFs exported
by Microsoft 365 for Mac 16.109 from the same checked-in source fixtures. The
gate verifies source and reference hashes, producer/version provenance, page
count and geometry, then compares 72-DPI rasters within recorded distance
budgets. These adapters remain `candidate`: an `exact` capability claim is
scoped to the named semantic invariant and does not mean whole-document pixel
equivalence. Static HTML uses a standards-oriented market corpus and approved
OfficeIMO regression baselines instead of pretending that one browser snapshot
defines HTML/CSS correctness.

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
| Create PDFs | Ready for common business documents | Fluent flow and canvas APIs cover text, links, lists, tables, mixed inline text/images/boxes, drawings, grouped header/footer text and images, watermarks, metadata, sections, generated TOCs, conditional/replayable flow, position capture, styled multipage containers, line-balanced multi-column flow, paragraph splitting with configurable widow/orphan counts, keep-with-next across block types, table tail-row control, generated optional-content layers, portfolios, form fields, tagging groundwork, and viewer settings. Shared Drawing contracts provide Unicode-safe line breaking, Latin ligatures, text direction, and host-provided shaping. | Built-in full complex-script shaping and deeper forms/annotations remain; continue unusual pagination and producer-specific visual fixtures as real failures are found. |
| Read and inspect | Ready for common born-digital PDFs | `PdfDocument.Open(...)` provides one bounded byte/path/stream source and reuses one canonical parse for text, geometry, images, attachments, portfolio metadata, outlines, links, annotations, forms, actions, metadata, XMP, tagged content, layers, output intents, security, revisions, signatures, diagnostics, and optional compliance readback. `PdfDocument.Preflight(...)` provides non-throwing text/security readiness for path, byte, and stream inputs before a full workflow is selected. `Analyze(...)` returns the consolidated health and capability report. `PdfReadOptions.Limits` bounds input bytes before buffering plus indirect objects, object characters/tokens/nesting, raw and decoded streams, content operations/operands/nesting, page counts, and page-tree depth/nodes. Strict mode rejects structural defects; lenient mode records explicit repairs. | Continue adding producer-specific repair fixtures; never auto-repair a defect whose semantic intent is ambiguous. |
| Merge PDFs | Ready for rewrite-safe inputs | `PdfDocument.MergeWith(...)` merges files, streams, bytes, or another opened document through the shared import/rewrite engine. `PdfDocument.Merge(...)` accepts a prepared document sequence and performs one shared merge pass for thin consumers. Pages can be normalized and supported visual annotations flattened. | Broader complex-file and producer interoperability proof. |
| Split and extract pages | Ready for rewrite-safe inputs | Single pages, page ranges, range expressions, fixed-size groups, and bookmark-derived ranges are supported. | Better preservation policy reporting for structures whose targets fall outside the selected pages. |
| Remove, duplicate, move, reorder, and rotate pages | Ready for rewrite-safe inputs | Fluent and static APIs cover the standard page-editing operations. `ComposePages`/`ComposePageRanges` allow selected subsets and repetitions through the shared extraction engine; convenience APIs reverse documents, repeat selections, and round-robin interleave even or uneven ranges. | Broader object-stream, tagged, layered, form-heavy, attachment-heavy, and incremental-file proof. |
| Copy pages from another PDF | Ready for rewrite-safe inputs | Pages can be appended, prepended, or inserted from another PDF, with optional annotation flattening. | The same collision and catalog policies needed by merge, plus a concise import report. |
| Resize pages | Ready for the supported rewrite subset | Pages can be resized with fit/fill/stretch behavior and destination transforms. | Broader preservation and visual corpus for inherited resources and unusual page trees. |
| Crop pages | Partial | Any production boundary box can be set, including `/CropBox`, `/TrimBox`, `/BleedBox`, `/ArtBox`, and `/MediaBox`. | Add named crop APIs, crop-and-translate, and an explicitly destructive crop mode that removes or clips content outside the retained area. Setting `/CropBox` alone must not be described as content removal. |
| Stamp and watermark | Ready for the supported rewrite subset | Text and image stamps/watermarks can target selected pages through fluent or static APIs. Complete source PDF pages can be imported as vector-preserving Form XObjects above or below selected target content with fit, alignment, rectangle, and opacity controls. | Rich text, reusable appearance templates, optional append-only stamping when signature permissions allow it, and wider resource-preservation proof. |
| Bookmarks and outlines | Partial | Generated documents can create nested outlines and named destinations; existing outlines can be read, preserved when supported, and used to split a document. | Add an existing-document outline editor: add, remove, rename, move, nest, retarget, rebuild from headings, and validate broken destinations. |
| Annotations | Partial | Generated PDFs can create text, free-text, highlight, and link annotations. Existing annotations can be read, filtered, flattened for supported appearances, updated in a small metadata/style subset, or removed. Page-to-image projection renders authored normal appearances and synthesizes a bounded appearance for supported free-text, text-markup, shape, line, ink, path, stamp, and caret annotations when `/AP` is missing; the synthesized case remains an explicit approximation diagnostic. Updates and removals can use append-only revisions for unsigned/approval-signed inputs and certification signatures with DocMDP `/P 3`; `/P 1` and `/P 2` are blocked, and widget edits remain routed through the FieldMDP-aware form engine. Results expose the mutation plan plus rewrite-preservation or signature/revision proof. | Add annotations to existing pages; move/resize them; edit subtype-specific geometry and appearance; reply/thread support; selective flattening; broader file-attachment and redaction annotation behavior. |
| Password protection | Ready for supported Standard-security workflows | Generated and rewritten PDFs default to AES-256 revision 6, with AES-128 interoperability and explicit legacy RC4 modes, typed permissions, Unicode password handling, revision 2-6 reading, authenticated user/owner roles, and owner-authorized encrypt/decrypt/re-encrypt workflows with preservation reports. | Expand encrypted mutation coverage beyond the dedicated security rewrite and keep signed/security-sensitive inputs fail-closed. Certificate signing and validation use the shared `OfficeIMO.Security` owner. |
| Metadata | Partial | Info-dictionary title, author, subject, and keywords can be replaced or updated by full rewrite or append-only revision. Generated PDFs can emit XMP and profile metadata; existing XMP is readable. | Edit and synchronize Info plus XMP, preserve custom schemas, manage dates/producer/creator deliberately, and report conflicts instead of silently choosing one source. |
| Forms | Broad | Generated and existing-document AcroForms support field creation, rename, remove, move, defaults, flags, calculation and tab order, appearance regeneration, exact-field flattening, typed/XFDF data interchange, append-only value updates, and empty signature-field placement. XFA is detected and explicitly rejected by the AcroForm editor rather than executed or silently changed. | Expand field kinds and appearance fidelity only when backed by interoperable fixtures; keep XFA outside the dependency-light core. |
| Incremental updates | Partial but real | A shared incremental object writer appends metadata, supported form values/appearances, external-signature preparation, and DSS/VRI validation material without replacing prior bytes. It preserves object generations and trailer state and emits classic or xref-stream revisions. Mutation plans and before/after reports prove byte-prefix, revision-chain, signature-range, and DocMDP/FieldMDP state. | Add encrypted incremental serialization, supported annotation/stamp/catalog operations, and broader interoperability fixtures. |
| Digital signatures | Partial | Approval, certification/DocMDP, and document-timestamp profiles can prepare external signatures; approval/certification fields can have visible widget appearances. `PdfCmsExternalSigner` and `PdfCmsSignatureCryptographyProvider` are thin adapters over the shared `OfficeIMO.Security` engine, which provides bounded CMS/RFC 3161 processing, RSA and ECDSA verification, platform X.509 chain/revocation policy, and one Bouncy Castle dependency across PDF and Email. Platform RSA signing does not export the private key. After signature math and digest verification, `PdfLongTermValidationEnricher` can append DER certificate, OCSP, and CRL streams in an ETSI DSS/VRI revision while retaining all earlier bytes and evidence. Reports keep structure, math, digest, trust, revocation, time, permissions, and later revisions separate. | Add deeper timestamp/revocation and external interoperability fixtures for B-LT/B-LTA workflows without claiming conformance prematurely. |
| Attachments and portfolios | Broad | Generated associated/embedded files and collection dictionaries are supported, including portfolio fields, sort order, initial document, and view. Existing attachments can be listed, selectively extracted, added, replaced, renamed, removed, and edited through the attachment engine; portfolio metadata is readable and supported rewrites retain it when preservation proof passes. | Add a focused existing-portfolio metadata editor and broaden viewer/interoperability fixtures for collection presentation. |
| Optimization | Broad | Deterministic Balanced, MaximumCompression, Web, Archival, and Custom profiles support lossless stream compression, unreachable-object removal, exact-stream and decoded-image deduplication, font/resource dictionary deduplication, classic or xref-stream output, object-stream packing, keep-original-if-larger behavior, per-action reporting, and post-save preservation proof. The Web profile emits standards-compliant Fast Web View output with two cross-reference sections plus page-offset and shared-object hint tables; linearization deliberately requires classic cross-reference tables without object streams. | Expand semantic deduplication and linearization corpus coverage only with bounded decoders and interoperable fixtures; optimization remains an explicit full rewrite and never claims signature preservation. |
| Redaction | Secure workflow available | Reviewable geometry/search plans remove intersecting text, vector paths, annotations, form fields, and image pixels. Built-in image normalization covers transformed placements, indexed/color-key/explicit/soft masks, and clone-on-write reuse; JPEG and other codecs use an optional bounded decoder contract or an explicit fail-closed/whole-placement policy. Cleanup policies cover metadata, attachments, structure/alternate text, and optional content. Proof combines extraction, raw/decoded residue checks, managed rendering, and pluggable independent validators. | Expand the hostile/corpus fixture set as new producer-specific encodings are found. |
| Render PDF pages | Broad managed subset with explicit diagnostics | Static pages project to shared Drawing primitives with paths/clipping, forms, images, axial/radial shadings, colored and basic uncolored vector tiling fills, supported annotation/form appearances, alpha, standard blend modes, Form-XObject alpha/luminosity soft masks, exact embedded TrueType outlines, and managed CMYK/Lab plus simplified calibrated-gray/RGB conversion through `OfficeIMO.Drawing`. The shared Drawing raster and SVG paths own reusable tiling, blending, and masking. PNG/SVG batches provide ranges, DPI/scale/background, thumbnails, cancellation, budgets, and per-page reports. A generated manifest reports simplified/unsupported operators and resources, and optional image codecs plug into shared Drawing rasterization without becoming core dependencies. | Extend fidelity from corpus failures while keeping CFF/Type 3 gaps, resource-specific calibrated parameters, unsupported ICC spaces, stroked/text tiling patterns and other broader pattern edge cases, and incomplete layer projection explicit in page reports. |
| Serialize generated PDFs | Bounded payload streaming | `PdfOptions.PageContentMemoryLimitBytes` bounds completed page/effect content retained during layout, and `PdfOptions.ObjectBufferMemoryLimitBytes` bounds completed indirect-object bytes during serialization. Both stores spill excess payloads to indexed temporary files; large stream objects are spooled without a duplicate combined buffer, and final stream assembly copies spilled objects in bounded chunks for plain and encrypted saves. Spill files are removed on disposal. | Per-page metadata and the authored block model remain proportional to document size, the active page is materialized while it is processed, and `ToBytes()` necessarily buffers the final artifact. Fully forward-only layout/output needs a deeper writer contract and representative memory gates. |
| Text and layout extraction | Broad, strategy-driven | The fast heuristic remains the default. A pluggable six-stage understanding pipeline provides confidence/evidence and stable JSON, Markdown, ALTO, hOCR, and PAGE XML. The built-in advanced profile adds rotation/arbitrary-baseline grouping, spatial and non-rectangular regions, multi-column/spanning-band order, tables, captions, headers/footers, and footnotes. | Refine advanced heuristics from real mixed-layout corpora and use provider stages for domain-specific reconstruction rather than hard-coding every document family. |
| PDF to Office/HTML/data | Partial by design | PDF-to-HTML review output, table export, Reader chunks, and limited PowerPoint table import use the shared logical model. | Improve the logical model and confidence/proof first. Do not promise general editable reconstruction from a presentation format. |
| Office/OpenDocument/HTML/Markdown/RTF/OneNote/AsciiDoc/LaTeX to PDF | Broad but evolving | Thin adapters use the shared PDF and Drawing engines. Word and PowerPoint preserve source font families and richer table/list/header/footer geometry; Excel uses a worksheet scene with authored row/column geometry, print areas, titles, breaks, charts, images, and conditional formatting; OpenDocument text, spreadsheet, and presentation formats expose one direct loss-aware façade over their existing semantic and PDF engines; static HTML uses the shared paged render scene with market-corpus raster gates, tables, forms, word breaking, and searchable text. OneNote, AsciiDoc, and LaTeX use explicit loss-aware semantic projections with combined diagnostics. | Browser-executed HTML is outside the current scope. Continue converter-specific fidelity only when the missing primitive is truly source-specific; otherwise improve the shared PDF, Drawing, HTML, or semantic-projection owner. |
| PDF/A, PDF/UA, and e-invoices | Exact-artifact proof available for declared profiles | PDF/A-2b, PDF/A-3b, PDF/UA-1, Factur-X, and ZUGFeRD generation gates combine internal readiness with external validator evidence bound to validator name/version/profile, SHA-256, byte length, result, and validation time. A report cannot be claimable while an effective requirement remains missing or unsupported. | Keep validator versions and profile fixtures current; do not broaden claims beyond exact artifacts that pass both internal and external proof. |

## Conversion Direction And Fidelity Assessment

| Direction | Assessment | Quality contract |
| --- | --- | --- |
| Office, OpenDocument, HTML, Markdown, RTF, OneNote, AsciiDoc, and LaTeX to PDF | Broad, with source-specific approximations | Every direct adapter uses the shared PDF/Drawing owners and returns stable conversion evidence. The generated support matrix distinguishes regression-proven, candidate, externally verified, and accepted-degradation routes instead of treating every feature as exact. |
| PDF to editable Word | Useful semantic recovery, not fixed-layout reconstruction | Metadata, page breaks, headings, paragraphs, lists, logical tables, links, supported images, and form placeholders are recovered when represented by the logical model. Unsupported image streams and interactive or unresolved navigation remain diagnostic-driven. |
| PDF to editable Excel or PowerPoint | Intentionally narrow | Logical tables can be recovered into worksheets or table slides with page/range limits and loss reports. Unrelated page text, drawings, images, and fixed layout are not advertised as editable reconstruction. |
| PDF to HTML | Good for semantic access and positioned visual review | Semantic and positioned-review profiles share the PDF logical/read model. Output is a review projection, not a browser-based reverse authoring guarantee. |
| Authored/loaded PDF to PNG, JPEG, TIFF, WebP, or SVG | Broad managed rendering with explicit gaps | One page-to-Drawing projection serves authored documents, loaded pages, batches, and source-conversion results with budgets, cancellation, selection, and diagnostics. Unsupported Type 3/CFF, ICC, pattern, and layer cases remain visible in page reports. |
| JPEG, PNG, GIF, BMP, TIFF, or supported WebP into authored/stamped PDF | Consistent shared ingestion | JPEG and writer-safe PNG embed directly. Other Drawing-decoded raster payloads normalize once to density-preserving PNG before all flow, table, inline, header/footer, background, watermark, canvas, and stamp paths. Malformed PNGs retain precise fail-closed diagnostics. |

## Remaining Engine Work, Easiest To Most Complex

The core is already broad enough for common business-document authoring,
inspection, conversion, rendering, and controlled mutation. Premium claims
remain capability-scoped and evidence-backed. The remaining engine work is
ordered by expected implementation and proof complexity:

1. Add a small set of tested report, invoice, label, and ticket recipes as
   `IPdfComponent` implementations after their contracts are stable. They must
   remain examples over the existing flow engine, not a template subsystem.
2. Broaden Drawing-owned raster ingestion where real fixtures require it,
   including explicit frame-selection and animation-loss policy. PDF surfaces
   should continue to consume one normalized payload contract.
3. Promote email, EPUB, and Visio routes only after body/attachment,
   book-resource/pagination, and vector-page policies are explicit. Each route
   should merge source and PDF evidence through the existing result contract.
4. Deepen generated and existing-document annotations, navigation, form-field
   kinds, and appearance editing through the current mutation and incremental
   update engines.
5. Add a dependency-light built-in complex-script shaping implementation, or a
   narrowly packaged provider, while retaining Drawing as the single shaping
   contract owner and preserving font fallback/subsetting proof.
6. Extend stateful and page-dependent composition for workflows whose content
   changes after pagination, without creating a second measurement or layout
   path for components.
7. Expand producer interoperability for Type 3/CFF rendering, ICC and pattern
   edge cases, layers, encrypted incremental updates, and complex structure
   preservation from corpus failures with explicit diagnostics.
8. Redesign layout and serialization for genuinely forward-only output with
   bounded memory, replay rules, deterministic object allocation, and artifact
   proof. This is architectural work; asynchronous save alone is not a
   streaming-layout guarantee.

## Current Architecture To Keep

- `PdfDocument.Create(...)` is the normal document-authoring entry point.
- `PdfDocument.Open(...)` is the normal fluent read and processing entry point.
- `PdfDocument.Read`, `Pages`, `Forms`, `Attachments`, `Bookmarks`,
  `Annotations`, and `Stamp` are the public workflow surfaces. The static
  parsing, inspection, manipulation, rendering, diagnostics, compliance, and
  signature engines are implementation owners rather than a second public API.
- `PdfReadDocument` is the canonical parser/read model. One opened
  `PdfDocument` snapshots its source once and reuses that parse across
  operations.
- `PdfDocument.Analyze(...)` is the consolidated health, capability,
  diagnostics, optimization, signature, repair, mutation, and optional
  compliance report.
- `PdfDocument.CreateComplianceArtifact(...)` captures exact output bytes and
  the matching writer/readback readiness in one immutable snapshot. External
  validators consume those bytes, and the snapshot reconciles their results
  without rerendering or accepting evidence for another artifact.
- `OfficeIMO.Drawing` is the shared managed scene, SVG, raster, Unicode
  line-breaking, Latin-ligature, text-direction, and host-shaping-contract owner.
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
  output intents, active content, and complex forms. Keep a small authoritative
  external gate hash-pinned to exact Open Preservation Foundation and veraPDF
  commits, with source paths, licenses, byte lengths, and focused expected
  behavior.
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
- [x] Keep PDF byte ranges, signature fields, revisions, and result mapping in
  `OfficeIMO.Pdf`, while the neutral `OfficeIMO.Security` owner provides bounded
  DER/CMS, signature math, certificate-chain, revocation, and RFC 3161 behavior
  through one Bouncy Castle dependency shared with Email.
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
- [x] Move text direction, shaped-run models, and the
  `IOfficeTextShapingProvider` seam into `OfficeIMO.Drawing` so PDF and future
  renderers share one host-integration contract. The dependency-free PDF engine
  continues to provide Unicode-scalar and Latin-ligature modes with explicit
  diagnostics; a host can adapt its chosen shaping engine without adding a
  dependency to the core packages.
- [x] Add shared block-aware multi-column flow, styled one-page containers,
  conditional/replayable flow constraints, position capture, semantic sections,
  generated TOCs, and optional-content layers before adapter-specific variants.
- [x] Add line-balanced columns, configurable paragraph widow/orphan counts,
  keep-with-next across block types, table tail-row control, and multipage
  decorated containers through the shared block-flow engine.
- [x] Add a machine-readable cross-format fidelity corpus and generated support
  matrix with source artifacts, hashes, text/logical proof, strict Poppler
  baselines, and explicit accepted-degradation policies.
- [x] Remove the three-family ceiling for embeddable generated text by adding
  shared named-font resources and carrying family selection through Word,
  Excel, PowerPoint, HTML, headings, tables, lists, headers, and footers.
- [x] Move Excel PDF fidelity to an authored worksheet scene and expand the
  static HTML renderer with a reusable market corpus, paged visual baselines,
  form controls, table layout, and word-breaking coverage.
- [x] Pin Microsoft Word, Excel, and PowerPoint 16.109 reference exports with
  source/reference hashes, producer provenance, page geometry, recorded visual
  distance results, and runtime-independent comparison gates.
- [ ] Keep HTML/CSS fidelity work in the canonical HTML/PDF/image plan. The PDF
  package should only add missing PDF-native primitives or writer support.
- [x] Add a reusable typed component contract that composes through the
  canonical flow engine, including existing layout constraints and position
  capture, without introducing another layout language.
- [ ] Add built-in report recipes only after repeated real examples identify a
  stable abstraction. Do not introduce a second template language or
  dependency-injection requirement into the PDF core speculatively.
- [ ] If real invoice, label, ticket, or logistics workflows require barcodes or
  QR codes, implement the encoding and drawing once in the smallest reusable
  `OfficeIMO.Drawing` owner (or a narrow optional adapter). PDF, Word, Excel,
  PowerPoint, and HTML surfaces should place the same typed drawing primitive.
- [ ] Expand generated annotations, form fields, signature appearances,
  attachments, layers, navigation, tagged structure, and accessible names
  through the same primitives used by existing-document editing.
- [ ] Continue Word, Excel, PowerPoint, Markdown, HTML, RTF, OneNote, AsciiDoc,
  and LaTeX fidelity through shared primitives or their established semantic
  projections first, with conversion warnings and proof snapshots preserved
  through post-processing.
- [x] Promote AsciiDoc and LaTeX compositions to direct adapters whose native
  parser and semantic-projection diagnostics flow automatically into the final
  `PdfDocumentConversionResult` while reusing `OfficeIMO.Markdown.Pdf`.
- [x] Promote OpenDocument text, spreadsheet, and presentation compositions to
  one direct adapter whose existing `OdfConversionResult` diagnostics flow
  automatically into the final PDF result without duplicating the Word, Excel,
  or PowerPoint PDF engines. Email, EPUB, and Visio remain planned until their
  body/asset/book/vector-page policies are explicit and visually proven.

Exit criterion: new fidelity work improves the shared engine or has a documented
source-format reason to remain in a thin adapter.

### P3 - Earn Narrow Compliance Claims

- [x] Gate PDF/A-2b and PDF/A-3b claims on exact-artifact veraPDF evidence
  without hiding internal readiness warnings.
- [x] Gate PDF/UA-1 on structure, reading order, alternate text, annotations,
  forms, links, language, fonts, Unicode mapping, and external validation.
- [x] Validate Factur-X/ZUGFeRD attachment, XMP, output-intent, invoice rules,
  and external PDF/invoice validator results as one artifact proof.
- [x] Store validator name, version, profile, result, validation time, artifact
  SHA-256, and byte length; reconcile externally satisfied requirements so a
  claimable export cannot also report missing or unsupported requirements.

Exit criterion: a profile is advertised only when internal readiness and the
external validator both pass the exact generated artifact.

### P4 - Productize The Library Surface

- [x] Route byte, stream, path, sync, async, and fluent opening through one
  bounded immutable source and one canonical parser without multiplying
  independent implementations.
- [x] Preserve source/output warnings through chained operations and expose
  exact artifact hashes, byte and page counts, operation names, observed
  mutation execution modes, timings, and failures. `Save(...)`,
  `SaveAsync(...)`, `TrySave(...)`, and typed adapter `SaveAsPdf(...)` routes
  return the same `PdfSaveResult` shape with a shared immutable pipeline report.
- [x] Bound completed page/effect content and serialized-object retention for
  stream saves with independent memory thresholds, indexed temporary-file
  spillover, direct large-stream spooling, and chunked final assembly, including
  encrypted output.
- [ ] Add fully forward-only layout and serialization before describing async
  APIs as fully streaming. Per-page metadata and the authored block model remain
  proportional to document size, while `ToBytes()` buffers the final artifact.
- [x] Generate and CI-check the public PDF conversion support matrix from the
  tested capability manifest so route and diagnostic claims cannot drift.
- [ ] Generate README examples from tested capability records where doing so
  improves user guidance without making prose less readable.
- [x] Keep a public-surface/dependency contract plus dependency-free mixed
  60-page cold/cached analysis, SVG, and PNG performance budgets in CI; verify
  output integrity as well as time and allocation ceilings, and keep bounded
  hostile-input, cross-platform, deterministic-output, and visual contracts in
  the focused suites.
- [x] Isolate optional dependencies in narrow adapter packages. The core
  exposes provider contracts for cryptography, OCR, advanced shaping, or codecs
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
