# OfficeIMO OneNote current state

OfficeIMO owns the offline OneNote format engine. Microsoft Graph and GraphEssentialsX are outside this implementation boundary; they can be added later as optional cloud transport without becoming a prerequisite for local files.

## Artifact coverage

| Capability | Desktop `.one` | FSSHTTP `.one` | `.onetoc2` | `.onepkg` |
| --- | :---: | :---: | :---: | :---: |
| Detect and validate | Yes | Yes | Yes | Yes |
| Typed read | Yes | Yes | Notebook hierarchy | Notebook hierarchy and sections |
| Create/write | Yes | Yes | Yes | Yes |
| Read-edit-write | Yes | Yes | Yes | Yes |
| Bounded parsing | Yes | Yes | Yes | Yes |
| Unknown-data preservation | Yes | Yes | Yes | Through contained native artifacts |

The two `.one` encodings share one semantic model. Conversion and Reader packages do not parse binary OneNote data themselves.

Loaded `.one` and `.onetoc2` artifacts preserve their desktop or FSSHTTP physical encoding by default. `OneNoteWriterOptions.StorageFormat` can select either encoding explicitly; new native artifacts default to the desktop revision store. `.onepkg` remains a separate Cabinet container created through `OneNotePackageWriter`.

## Content fidelity

| Content | Read | New write/edit | Preservation behavior |
| --- | :---: | :---: | --- |
| Notebook/section/page/subpage hierarchy | Yes | Yes | Typed |
| Page size, orientation, margins, RTL/read-only flags, outlines, and collision-aware layout | Yes | Yes | Typed plus unknown properties |
| Rich text and styles | Yes | Yes | Typed plus unknown run properties |
| Lists, tables, and hyperlinks | Yes | Yes | Typed |
| Images, printouts, backgrounds, OCR, and embedded files | Yes | Yes | Lazy bounded payloads; primary and `WebPictureContainer14` relationships; unresolved loaded relationships retained during preservation writes |
| Note tags and task tags | Yes | Yes | Typed definitions and state |
| Authors, timestamps, metadata, and revisions | Yes | Yes | Typed plus opaque revision data |
| Conflict pages | Yes | Yes | Native child object spaces |
| Version-history pages | Yes | Yes | Native revision contexts |
| Ink/handwriting | Yes for native X/Y/pressure strokes and recognition | Yes | Shared editable strokes, pen style, transforms, LCID, recognized text, and alternatives; native scaling, nested containers, unknown packet dimensions, and opaque strokes retained without unsafe flattening |
| Structured math | Yes | Yes | Shared editable AST with native OneNote mapping for groups, delimiter lists, left/right scripts, limits, slashed fractions, stacks, matrices, arrays, and decorations plus bounded MathML, LaTeX, plain-text, and Drawing projections |
| Unknown objects/properties/relationships | Opaque | Not directly authored | Retained in sections and root/nested TOCs unless a typed edit replaces the owning relationship |

Picture width and height are IEEE-754 floating-point properties expressed in half-inch units, as defined by MS-ONE. The native model exposes those units directly; Reader reports 96-DPI pixel metadata. When an image has both native payload relationships, the normal picture container wins and `WebPictureContainer14` is used as a fallback. Unresolved payloads remain visible as metadata and retain their native relationships in preservation-mode writes.

OneNote/RichEdit vertical tabs and form feeds project as line breaks. Other unsupported control characters, Unicode noncharacters, and unpaired surrogate code units project as `?`. This normalization is limited to conversion and Reader output; it does not rewrite source strings in the typed OneNote model.

Current pages are the default Reader and conversion surface. `OneNoteMarkdownOptions` and `ReaderOneNoteOptions` provide separate opt-ins for conflict pages and version-history snapshots so applications do not accidentally ingest superseded content. Reader metadata still reports current, conflict, and version counts. Notebook readers similarly exclude `OneNote_RecycleBin` by default and expose it through `OneNoteNotebookReaderOptions.IncludeRecycleBin`.

Direct page content is accepted on read. Writers canonicalize it into an outline, preserving element content while producing the interoperable native hierarchy. Shared native page-series objects retain contiguous current-page runs and their ordered cached metadata during preservation writes; an insertion or move that splits a source series starts a new native run so caller order remains authoritative. Empty recycle-bin groups are written with a valid empty TOC.

Native handwriting recognition is read and written as the page-level root → line → block → word graph used by desktop OneNote. Stroke references are resolved in each recognition word's native object namespace. The public model keeps the reusable result on `OfficeInkStroke` instead of exposing OneNote-specific object identifiers.

Native recording media retains the recording GUID and unsigned millisecond duration, and each page writes its `AudioRecordingGuids` index from the media it contains. The writer accepts the MS-ONE recording extensions (`.wma`, `.mp3`, `.wav`, `.wmv`, `.avi`, and `.mpg`), infers audio/video kind when possible, and rejects mismatched or non-representable media rather than returning it later as a different typed element.

## Interoperability proof

The automated suite validates desktop revision-store structures, transaction checksums, read-only declaration hashes, dependency graphs, FSSHTTP stream objects and cells, shared multi-page series (including caller insertions), notebook TOCs, Cabinet packages, and semantic read-after-write behavior. The default readback guard checks page identity/order/relationships, structural content and table topology, rich-text runs, supported layout/media metadata, editable math and ink, and binary payload resolution plus known length; it is not a byte-for-byte payload or opaque-property equivalence check. A legal fixture corpus covers desktop, Microsoft 365/FSSHTTP, and real handwriting-recognition sources.

Manual desktop validation used Microsoft OneNote only as an interoperability oracle: generated sections containing rich text, tags/tasks, conflict pages, and version history were opened, edited, saved, closed, and reopened. OfficeIMO then read the OneNote-saved artifacts with no parser diagnostics and observed the external edits. No COM or OneNote dependency exists in the shipping libraries.

## Safety model

Parsing is bounded by configurable byte, node, transaction, object, property, recursion, stream-object, asset, package-entry, and expansion limits. Native and portable math parsing has an explicit 128-level default depth ceiling. Writing validates caller-created notebook, page, content, and math graphs for cycles, shared instances, excessive depth, OneNote's 255-column native matrix/array descriptor capacity, single-code-unit native math characters, canonical named/custom page geometry, and rectangular native table topology/widths. Opaque direct ink retains its complete source bounding box and unions authored strokes or fails closed. Shared Drawing image decoders cap encoded payloads and source pixels, reject overflowing PNG chunk ranges before allocation, and bound PNG inflation to its declared scanlines. Shared Markdown projection clamps arbitrary list levels before indentation reaches Reader, HTML, or PDF. Package entry names reject traversal, rooted paths, and drive paths. Deterministic byte-mutation and truncation tests require malformed inputs either to parse safely or fail through bounded I/O/format exceptions rather than runtime index or allocation failures.

## Projection ownership

```text
OfficeIMO.OneNote
    -> OfficeIMO.Drawing canvas
        -> PNG / JPEG / TIFF / SVG / WebP
        -> OfficeIMO.OneNote.Html visual pages
        -> OfficeIMO.OneNote.Pdf visual pages
    -> OfficeIMO.OneNote.Markdown semantic projection
        -> OfficeIMO.Reader.OneNote
        -> OfficeIMO.OneNote.Html semantic document
        -> OfficeIMO.OneNote.Pdf semantic document
```

This keeps one native parser, one semantic projection, and one positioned visual canvas. Ink and math are document-agnostic models in `OfficeIMO.Drawing`; OneNote owns only the native adapters. Word uses the same math tree through its OMML adapter, so neither document package depends on the other.

Visual HTML embeds one responsive SVG canvas per page and can include an encoded assistive-text projection. Semantic HTML reflows the Markdown model into normal document markup. Visual PDF places a full-page PNG generated from the shared canvas at the native page geometry; its default raster scale is 2 (144 DPI), with a configurable per-page pixel limit. The same Drawing-owned ceiling protects direct PNG/JPEG/TIFF/WebP export. Source images outside Drawing's built-in PNG/JPEG/BMP/GIF decoders can use the reusable `IOfficeRasterImageCodec` boundary; otherwise raster outputs contain a visible placeholder and a structured warning. Semantic PDF keeps selectable text and normal document reflow. Applications choose between layout preservation and semantic text behavior rather than receiving a silent hybrid.

OneNote PDF conversion opts into the shared multilingual system-font fallback unless `PdfTextFallbackFeatures.None` is selected explicitly. This improves coverage for valid CJK, Arabic, and other scripts when suitable fonts are installed; glyph availability remains host-dependent and unresolved valid glyphs are reported through PDF conversion diagnostics.

## Specifications

- [MS-ONE: OneNote File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-one/73d22548-a613-4350-8c23-07d15576be50)
- [MS-ONESTORE: OneNote Revision Store File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-onestore/)
- [MS-FSSHTTPB: Binary Requests for File Synchronization](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/)

The implementation is based on published format contracts and independently written MIT-licensed code. Fixture provenance is recorded in `OfficeIMO.OneNote.Tests/Fixtures/SOURCE.md`.
