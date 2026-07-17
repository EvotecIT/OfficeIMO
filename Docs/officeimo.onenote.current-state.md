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
| Outlines and layout | Yes | Yes | Typed plus unknown properties |
| Rich text and styles | Yes | Yes | Typed plus unknown run properties |
| Lists, tables, and hyperlinks | Yes | Yes | Typed |
| Images and embedded files | Yes | Yes | Lazy bounded payloads; primary and `WebPictureContainer14` relationships; unresolved loaded relationships retained during preservation writes |
| Note tags and task tags | Yes | Yes | Typed definitions and state |
| Authors, timestamps, metadata, and revisions | Yes | Yes | Typed plus opaque revision data |
| Conflict pages | Yes | Yes | Native child object spaces |
| Version-history pages | Yes | Yes | Native revision contexts |
| Ink/handwriting | Where decoded or available as payload | Not yet | Preserved during unrelated edits; replacement fails closed |
| Plain math | Yes | Yes | Typed |
| Raw MathML/LaTeX math payloads | Where present | Not yet | Preserved during unrelated edits; replacement fails closed |
| Unknown objects/properties/relationships | Opaque | Not directly authored | Retained in sections and root/nested TOCs unless a typed edit replaces the owning relationship |

Picture width and height are IEEE-754 floating-point properties expressed in half-inch units, as defined by MS-ONE. The native model exposes those units directly; Reader reports 96-DPI pixel metadata. When an image has both native payload relationships, the normal picture container wins and `WebPictureContainer14` is used as a fallback. Unresolved payloads remain visible as metadata and retain their native relationships in preservation-mode writes.

OneNote/RichEdit vertical tabs and form feeds project as line breaks. Other unsupported control characters, Unicode noncharacters, and unpaired surrogate code units project as `?`. This normalization is limited to conversion and Reader output; it does not rewrite source strings in the typed OneNote model.

Current pages are the default Reader and conversion surface. `OneNoteMarkdownOptions` and `ReaderOneNoteOptions` provide separate opt-ins for conflict pages and version-history snapshots so applications do not accidentally ingest superseded content. Reader metadata still reports current, conflict, and version counts. Notebook readers similarly exclude `OneNote_RecycleBin` by default and expose it through `OneNoteNotebookReaderOptions.IncludeRecycleBin`.

Direct page content is accepted on read. Writers canonicalize it into an outline, preserving element content while producing the interoperable native hierarchy. Shared native page-series objects retain contiguous current-page runs and their ordered cached metadata during preservation writes; an insertion or move that splits a source series starts a new native run so caller order remains authoritative. Empty recycle-bin groups are written with a valid empty TOC.

## Interoperability proof

The automated suite validates desktop revision-store structures, transaction checksums, read-only declaration hashes, dependency graphs, FSSHTTP stream objects and cells, shared multi-page series (including caller insertions), notebook TOCs, Cabinet packages, and semantic read-after-write behavior. The default readback guard checks page identity/order/relationships, structural content and table topology, rich-text runs, supported layout/media metadata, and binary payload resolution plus known length; it is not a byte-for-byte payload or opaque-property equivalence check. A legal fixture corpus covers both desktop and Microsoft 365/FSSHTTP sources.

Manual desktop validation used Microsoft OneNote only as an interoperability oracle: generated sections containing rich text, tags/tasks, conflict pages, and version history were opened, edited, saved, closed, and reopened. OfficeIMO then read the OneNote-saved artifacts with no parser diagnostics and observed the external edits. No COM or OneNote dependency exists in the shipping libraries.

## Safety model

Parsing is bounded by configurable byte, node, transaction, object, property, recursion, stream-object, asset, package-entry, and expansion limits. Writing validates caller-created notebook, page, and content graphs for cycles and shared instances, enforces configurable depth limits with a hard safety ceiling, reserves each scope's required table-of-contents filename, and rejects native list levels outside the representable range. Shared Markdown projection clamps arbitrary list levels before indentation reaches Reader, HTML, or PDF. Package entry names reject traversal, rooted paths, and drive paths. Deterministic byte-mutation and truncation tests require malformed inputs either to parse safely or fail through bounded I/O/format exceptions rather than runtime index or allocation failures.

## Projection ownership

```text
OfficeIMO.OneNote
    -> OfficeIMO.OneNote.Markdown
        -> OfficeIMO.Reader.OneNote
        -> OfficeIMO.OneNote.Html
        -> OfficeIMO.OneNote.Pdf
```

This keeps one native parser and one semantic projection. HTML and PDF reuse the first-party Markdown, HTML, and PDF engines.

OneNote PDF conversion opts into the shared multilingual system-font fallback unless `PdfTextFallbackFeatures.None` is selected explicitly. This improves coverage for valid CJK, Arabic, and other scripts when suitable fonts are installed; glyph availability remains host-dependent and unresolved valid glyphs are reported through PDF conversion diagnostics.

## Specifications

- [MS-ONE: OneNote File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-one/73d22548-a613-4350-8c23-07d15576be50)
- [MS-ONESTORE: OneNote Revision Store File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-onestore/)
- [MS-FSSHTTPB: Binary Requests for File Synchronization](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/)

The implementation is based on published format contracts and independently written MIT-licensed code. Fixture provenance is recorded in `OfficeIMO.OneNote.Tests/Fixtures/SOURCE.md`.
