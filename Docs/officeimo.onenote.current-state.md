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

## Content fidelity

| Content | Read | New write/edit | Preservation behavior |
| --- | :---: | :---: | --- |
| Notebook/section/page/subpage hierarchy | Yes | Yes | Typed |
| Outlines and layout | Yes | Yes | Typed plus unknown properties |
| Rich text and styles | Yes | Yes | Typed plus unknown run properties |
| Lists, tables, and hyperlinks | Yes | Yes | Typed |
| Images and embedded files | Yes | Yes | Lazy bounded payloads |
| Note tags and task tags | Yes | Yes | Typed definitions and state |
| Authors, timestamps, metadata, and revisions | Yes | Yes | Typed plus opaque revision data |
| Conflict pages | Yes | Yes | Native child object spaces |
| Version-history pages | Yes | Yes | Native revision contexts |
| Ink/handwriting | Where decoded or available as payload | Not yet | Preserved during unrelated edits; replacement fails closed |
| Plain math | Yes | Yes | Typed |
| Raw MathML/LaTeX math payloads | Where present | Not yet | Preserved during unrelated edits; replacement fails closed |
| Unknown objects/properties/relationships | Opaque | Not directly authored | Retained unless a typed edit replaces the owning relationship |

## Interoperability proof

The automated suite validates desktop revision-store structures, transaction checksums, read-only declaration hashes, dependency graphs, FSSHTTP stream objects and cells, notebook TOCs, Cabinet packages, and read-after-write behavior. A legal fixture corpus covers both desktop and Microsoft 365/FSSHTTP sources.

Manual desktop validation used Microsoft OneNote only as an interoperability oracle: generated sections containing rich text, tags/tasks, conflict pages, and version history were opened, edited, saved, closed, and reopened. OfficeIMO then read the OneNote-saved artifacts with no parser diagnostics and observed the external edits. No COM or OneNote dependency exists in the shipping libraries.

## Safety model

Parsing is bounded by configurable byte, node, transaction, object, property, recursion, stream-object, asset, package-entry, and expansion limits. Package entry names reject traversal, rooted paths, and drive paths. Deterministic byte-mutation and truncation tests require malformed inputs either to parse safely or fail through bounded I/O/format exceptions rather than runtime index or allocation failures.

## Projection ownership

```text
OfficeIMO.OneNote
    -> OfficeIMO.OneNote.Markdown
        -> OfficeIMO.Reader.OneNote
        -> OfficeIMO.OneNote.Html
        -> OfficeIMO.OneNote.Pdf
```

This keeps one native parser and one semantic projection. HTML and PDF reuse the first-party Markdown, HTML, and PDF engines.

## Specifications

- [MS-ONE: OneNote File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-one/73d22548-a613-4350-8c23-07d15576be50)
- [MS-ONESTORE: OneNote Revision Store File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-onestore/)
- [MS-FSSHTTPB: Binary Requests for File Synchronization](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-fsshttpb/)

The implementation is based on published format contracts and independently written MIT-licensed code. Fixture provenance is recorded in `OfficeIMO.OneNote.Tests/Fixtures/SOURCE.md`.
