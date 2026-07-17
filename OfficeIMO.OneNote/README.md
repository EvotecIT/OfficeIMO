# OfficeIMO.OneNote

`OfficeIMO.OneNote` is a managed, cross-platform engine for offline Microsoft OneNote files. It reads, creates, edits, and writes native OneNote artifacts without Microsoft Graph, GraphEssentialsX, COM automation, a OneNote installation, or a commercial file-format library.

## Supported artifacts

| Artifact | Read | Create/write | Notes |
| --- | :---: | :---: | --- |
| Desktop `.one` revision store | Yes | Yes | Native MS-ONESTORE transaction, manifest, revision, and object graphs |
| FSSHTTP-encoded `.one` | Yes | Yes | Native MS-FSSHTTPB package-store cells and object streams |
| `.onetoc2` notebook table of contents | Yes | Yes | Section groups, sections, ordering, colors, and notebook metadata |
| `.onepkg` notebook export | Yes | Yes | Managed Cabinet package reader/writer with bounded expansion |
| Notebook directory | Yes | Yes | Root and nested `.onetoc2` files plus section files |

Both desktop and FSSHTTP `.one` inputs project into the same typed model. File-format logic stays in this package; Reader and conversion packages are thin consumers.

## Read, edit, and save a section

```csharp
using OfficeIMO.OneNote;

OneNoteSection section = OneNoteSectionReader.Read("Projects.one");
OneNotePage page = section.Pages[0];

var paragraph = new OneNoteParagraph();
paragraph.Runs.Add(new OneNoteTextRun { Text = "Added offline by OfficeIMO" });
page.DirectContent.Add(paragraph);

section.Save("Projects-updated.one");
```

New documents use the same model:

```csharp
var section = new OneNoteSection { Name = "Planning" };
var page = new OneNotePage { Title = "Release checklist" };
var item = new OneNoteParagraph {
    List = new OneNoteListInfo { Ordered = false, Level = 0 }
};
item.Runs.Add(new OneNoteTextRun { Text = "Validate the packaged artifact" });
page.DirectContent.Add(item);
section.Pages.Add(page);

OneNoteSectionWriter.Write(section, "Planning.one");
```

Writers validate their output by reading it back by default. A loaded `.one` or `.onetoc2` keeps its desktop or FSSHTTP physical encoding when saved; a new artifact defaults to the desktop revision store. Applications can select FSSHTTP output explicitly:

```csharp
var options = new OneNoteWriterOptions {
    StorageFormat = OneNoteStorageFormat.FileSynchronizationPackage
};

OneNoteSectionWriter.Write(section, "Planning.one", options);
```

`OneNoteWriterOptions.StorageFormat` applies to native `.one` and `.onetoc2` payloads. Use `OneNotePackageWriter` for the Cabinet-based `.onepkg` container. The same options also provide tighter output, package-entry, section-group-depth, related-page-depth, and content-depth limits when needed.

## Notebooks and packages

```csharp
var notebook = new OneNoteNotebook { Name = "Offline notebook" };
notebook.Sections.Add(section);

OneNoteNotebookWriter.Write(notebook, "Offline notebook"); // directory + .onetoc2 + .one
OneNotePackageWriter.Write(notebook, "Offline notebook.onepkg");

OneNoteNotebook reopened = OneNotePackageReader.Read("Offline notebook.onepkg");
```

`OneNoteNotebookReader` opens a notebook directory or `.onetoc2`. `OneNotePackageReader` opens `.onepkg`. Both retain section-group and page/subpage hierarchy.

## Semantic coverage

The typed model covers:

- notebooks, section groups, sections, pages, and subpages;
- positioned outlines, rich text and run styles, lists, hyperlinks, and tables;
- images, embedded files, recordings/media, lazy payloads, and layout metadata;
- note tags, Outlook-style task tags, authors, timestamps, and revisions;
- conflict copies and version-history pages;
- ink/handwriting payloads and decoded strokes where the representation is understood;
- plain and structured math projections where present;
- diagnostics plus unknown objects, properties, roots, and relationships for loss-aware preservation.

Native picture dimensions are exposed as `WidthHalfInches` and `HeightHalfInches`. MS-ONE stores these properties as IEEE-754 floating-point counts of half-inch units; Reader converts them to pixels at 96 DPI when it emits asset metadata. Images can carry both the normal `PictureContainer` relationship and the newer `WebPictureContainer14` relationship. Reading prefers the normal payload and falls back to the web payload when necessary. A loaded image whose payload cannot be resolved still retains its native relationships and can survive an unrelated preservation write; explicitly canonicalizing such an image without a payload fails instead of inventing or silently dropping data.

When a loaded section is edited, unsupported source structures are preserved unless the typed edit replaces or deletes their owning relationship. Known typed properties win over stale opaque values. Native `PageSeriesNode` objects can own several current pages; preservation writes keep contiguous members in that shared series shape and align cached page metadata by reference order instead of collapsing the series to one page. An insertion or move that splits the source series starts a new native series run, so preservation never overrides the caller's requested page order.

OneNote permits content directly below a page object. OfficeIMO reads that content through `OneNotePage.DirectContent`; on write it moves those elements into a canonical outline because that is the interoperable native shape. This is a structural normalization, not a content deletion.

New notebooks, groups, sections, pages, and content receive their native logical identities on first serialization. The assigned identities remain on the typed model and are reused by later saves; physical transaction and file-version identifiers still change for each artifact.

## Deliberate write boundaries

- New plain-text math can be serialized. Creating or replacing raw MathML/LaTeX payloads currently fails with `ONENOTE_WRITE_UNSUPPORTED_MATH` instead of flattening them silently.
- Source ink is retained during unrelated edits. Creating or replacing native ink currently fails with `ONENOTE_WRITE_UNSUPPORTED_INK` instead of dropping strokes.
- MS-ONE task tags are always checkable. A task or explicit normal-tag shape that contradicts `IsCheckable` fails closed instead of silently changing the tag after a round trip.
- Encrypted or otherwise unsupported sections produce diagnostics; notebook readers can continue with other sections when configured to do so.

These fail-closed boundaries distinguish preservation from authoring support.

## Safety and streaming

`OneNoteReaderOptions` bounds input bytes, file-node and transaction counts, objects, properties, property nesting, distinct page-graph nodes, related-page depth, assets, and FSSHTTP stream objects. Notebook/package options additionally bound section-group depth, entry count, per-entry bytes, and total expanded bytes. Lazy binary payloads are materialized only when requested and still require a caller-provided byte limit.

Conflict and version-history object spaces are traversed once per section. Repeated or cyclic references in malformed page graphs are pruned to a bounded spanning tree so the loaded model remains safe to convert and rewrite.

Before writing or projecting, caller-created section-group, conflict/version, and nested-content relationships are checked for cycles, shared instances, and excessive depth. `MaxSectionGroupDepth` defaults to 32; `MaxPageRelationshipDepth` and `MaxContentDepth` default to 128. Writer and Markdown options accept values up to the hard safety ceiling of 256, and configured writer bounds are propagated into read-back validation. Direct Markdown, HTML, and PDF conversion validates conflict/version branches only when requested; Reader validates them unconditionally because it reports their counts in metadata. The required `Open Notebook.onetoc2` filename is reserved when section-group directories are allocated. Native list levels are limited to 0 through `OneNoteListInfo.MaxLevel` (254); the shared Markdown projection clamps out-of-range caller values to that representable range so Markdown, HTML, PDF, and Reader conversion cannot allocate arbitrary indentation.

`OneNoteWriterOptions.ValidateRoundTrip` is enabled by default. Section writes, including sections emitted inside `.onepkg`, are read back before bytes are returned. Validation covers page identity, order, conflict/version relationships, titles and core page metadata; outline and table hierarchy; rich-text run boundaries, text, links, and supported formatting; supported layout, image, attachment, media, and math metadata; and binary payload resolution plus known length. This is a semantic guard against silent model loss, not a byte-for-byte payload hash or opaque-property equivalence check.

Caller-owned streams stay open. Seekable read streams are restored to their original position. Async probe and Reader entry points support cancellation.

## Conversion and Reader packages

- `OfficeIMO.OneNote.Markdown` owns the shared semantic Markdown/text projection.
- `OfficeIMO.OneNote.Html` renders HTML through the first-party Markdown model.
- `OfficeIMO.OneNote.Pdf` renders PDF through `OfficeIMO.Markdown.Pdf` and `OfficeIMO.Pdf`.
- `OfficeIMO.Reader.OneNote` emits page-aware chunks, hierarchy, tables, links, assets, metadata, hashes, diagnostics, and Markdown/text projections.

Conflict copies and version-history snapshots are opt-in in direct conversions through `OneNoteMarkdownOptions`. Reader reports their counts in structured metadata while keeping current pages as the default chunk surface.

Notebook readers exclude the `OneNote_RecycleBin` section group by default. Set `OneNoteNotebookReaderOptions.IncludeRecycleBin = true` when an application needs to inspect it. When writing, mark the group with `IsRecycleBin = true`; the writer emits the canonical directory name even when the model uses a different display name. An empty recycle-bin group is valid and is written with a valid empty table of contents.

## Compatibility evidence

The test corpus contains legally reusable Apache-2.0 OneNote fixtures for desktop and Microsoft 365/FSSHTTP encodings. Tests cover native read/write round trips, unknown-data preservation, deterministic corruption mutations, truncation, limits, package paths, and all supported target frameworks. Release validation also covers packed-NuGet consumer use. Generated desktop sections have also been opened, edited, saved, closed, and reopened with Microsoft OneNote during interoperability validation; OneNote is not required at runtime or in CI.

See the [current-state and capability matrix](../Docs/officeimo.onenote.current-state.md) for detailed boundaries and links to the Microsoft format specifications.
