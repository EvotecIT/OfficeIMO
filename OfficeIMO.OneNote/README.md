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

`OneNoteWriterOptions.StorageFormat` applies to native `.one` and `.onetoc2` payloads. Use `OneNotePackageWriter` for the Cabinet-based `.onepkg` container. The same options also provide tighter output and package-entry limits when needed.

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

When a loaded section is edited, unsupported source structures are preserved unless the typed edit replaces or deletes their owning relationship. Known typed properties win over stale opaque values.

New notebooks, groups, sections, pages, and content receive their native logical identities on first serialization. The assigned identities remain on the typed model and are reused by later saves; physical transaction and file-version identifiers still change for each artifact.

## Deliberate write boundaries

- New plain-text math can be serialized. Creating or replacing raw MathML/LaTeX payloads currently fails with `ONENOTE_WRITE_UNSUPPORTED_MATH` instead of flattening them silently.
- Source ink is retained during unrelated edits. Creating or replacing native ink currently fails with `ONENOTE_WRITE_UNSUPPORTED_INK` instead of dropping strokes.
- MS-ONE task tags are always checkable. A task or explicit normal-tag shape that contradicts `IsCheckable` fails closed instead of silently changing the tag after a round trip.
- Encrypted or otherwise unsupported sections produce diagnostics; notebook readers can continue with other sections when configured to do so.

These fail-closed boundaries distinguish preservation from authoring support.

## Safety and streaming

`OneNoteReaderOptions` bounds input bytes, file-node and transaction counts, objects, properties, nesting depth, assets, and FSSHTTP stream objects. Notebook/package options additionally bound section-group depth, entry count, per-entry bytes, and total expanded bytes. Lazy binary payloads are materialized only when requested and still require a caller-provided byte limit.

Caller-owned streams stay open. Seekable read streams are restored to their original position. Async probe and Reader entry points support cancellation.

## Conversion and Reader packages

- `OfficeIMO.OneNote.Markdown` owns the shared semantic Markdown/text projection.
- `OfficeIMO.OneNote.Html` renders HTML through the first-party Markdown model.
- `OfficeIMO.OneNote.Pdf` renders PDF through `OfficeIMO.Markdown.Pdf` and `OfficeIMO.Pdf`.
- `OfficeIMO.Reader.OneNote` emits page-aware chunks, hierarchy, tables, links, assets, metadata, hashes, diagnostics, and Markdown/text projections.

Conflict copies and version-history snapshots are opt-in in direct conversions through `OneNoteMarkdownOptions`. Reader reports their counts in structured metadata while keeping current pages as the default chunk surface.

## Compatibility evidence

The test corpus contains legally reusable Apache-2.0 OneNote fixtures for desktop and Microsoft 365/FSSHTTP encodings. Tests cover native read/write round trips, unknown-data preservation, deterministic corruption mutations, truncation, limits, package paths, and all supported target frameworks. Release validation also covers packed-NuGet consumer use. Generated desktop sections have also been opened, edited, saved, closed, and reopened with Microsoft OneNote during interoperability validation; OneNote is not required at runtime or in CI.

See the [current-state and capability matrix](../Docs/officeimo.onenote.current-state.md) for detailed boundaries and links to the Microsoft format specifications.
