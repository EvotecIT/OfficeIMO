# Hidden data and personal information support matrix

`WordDocument`, `ExcelDocument`, `PowerPointDocument`, `VisioDocument`, `OdfDocument` (ODT, ODS, ODP and ODG), `EpubPublication` and `PdfDocument` find and remove the metadata and hidden data that Office's Document Inspector covers. `InspectHiddenData()` changes nothing and returns an `OfficeHiddenDataReport`. `RemoveHiddenData(OfficeHiddenDataOptions)` removes the selected categories in memory and returns the same report shape, where each finding records what happened to it. Save the document to keep the result. `PdfDocument.RemoveHiddenData` returns an `OfficeHiddenDataResult` with the cleaned bytes instead, because a PDF is cleaned by one verified rewrite and the loaded document is not changed.

## Which call to make

| You have | To see what leaves with it | To clean it |
| --- | --- | --- |
| A loaded document | `document.InspectHiddenData()` | `document.RemoveHiddenData(options)`, then save; for a PDF, write `result.ToArray()` |
| A file | `OfficeDocument.InspectHiddenData(path)` | `OfficeDocument.RemoveHiddenData(inputPath, outputPath, options)` |
| Bytes, such as an upload in a browser | `OfficeDocument.InspectHiddenData(data, fileName)` | `OfficeDocument.RemoveHiddenData(data, fileName, options)` |
| A file, with content and summaries too | `OfficeDocument.Inspect(path, new OfficeInspectOptions { IncludeHiddenData = true })` | |
| A file, provenance only, without loading it | `OfficeDocument.Inspect(path, new OfficeInspectOptions { Provenance = new() })` | The format's static `RemoveProvenance(inputPath, outputPath)` |
| A shell or an agent | `officeimo hidden-data inspect` or the `officeimo_inspect_hidden_data` MCP tool | `officeimo hidden-data remove` or `officeimo_remove_hidden_data` |

Saving after removal follows the same signature rule everywhere: `OfficeSignatureMutationPolicy.BlockSave` (the default) refuses to invalidate a signature, `RemoveInvalidatedSignatures` removes the signatures, and `PreserveSignatureMarkup` keeps signature markup that no longer verifies. Word, Excel and PowerPoint take it as `SignatureMutationPolicy` on their save options, `OfficeDocument.RemoveHiddenData` as a parameter, and `RemoveProvenance` on `OfficeProvenanceRemovalOptions`.

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("Contract.docx");
OfficeHiddenDataReport report = document.RemoveHiddenData(new OfficeHiddenDataOptions {
    Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.HiddenContent,
    TrackedChanges = OfficeTrackedChangesAction.AcceptAll,
    AnonymousAuthor = "Reviewer"
});
foreach (OfficeHiddenDataFinding finding in report.Findings) {
    Console.WriteLine($"{finding.Outcome} {finding.Kind} at {finding.Location}: {finding.Description}");
}
document.Save("Contract-shared.docx");
```

`OfficeHiddenDataKind.Default` selects every category except tracked changes and hidden content, which change what readers see, attachments and bookmarks, which carry or organize the content being shared, and provenance, which discloses how content was made. Active content such as scripts is part of `Default`. Tracked changes are reported and kept unless `TrackedChanges` is `AcceptAll` or `RejectAll`. Authors on comments and revisions that stay in the document are replaced with `AnonymousAuthor` (default `Author`) rather than deleted, because Office requires an author name on them.

## Provenance

The report also answers what provenance leaves with the file. Content Credentials (C2PA manifests and references to them) and IPTC digital source type declarations, such as an image marked as AI-generated, appear as `Provenance` findings whose `Provenance` property carries the same evidence that `InspectProvenance` returns.

Provenance is kept unless you add `OfficeHiddenDataKind.Provenance`. While it is kept, `ImageMetadata` cleanup reduces an image's XMP to its digital source type declaration instead of deleting it, and leaves images that carry Content Credentials untouched, because changing their bytes would invalidate the credential; those images are reported as `Retained`. Selecting `Provenance` removes the carriers through the same engine as `RemoveProvenance`, and then image metadata is stripped as usual. `OfficeHiddenDataOptions.Provenance` chooses the carriers with the same switches as `RemoveProvenance`: to keep Content Credentials but drop AI-generated declarations, set `RemoveC2paManifests` to false. Removing a declaration changes the image, so a kept credential is reported with a note that it no longer verifies. Each finding's `Scope` says whether it belongs to the document, one embedded image, or the whole file.

Content Credentials stored for the whole package (`META-INF/content_credential.c2pa`) cover every byte of the file, so any change to a loaded document invalidates them. `RemoveHiddenData` always removes them and says so in the finding's `Note`. To remove only provenance from a file and leave every other byte alone, use the static `RemoveProvenance(inputPath, outputPath)` instead; it also handles package signatures through `SignatureMutationPolicy`.

## Outcomes

Each finding's `Outcome` is `Found` when inspecting, and `Removed`, `Anonymized`, `Accepted`, `Rejected` or `Retained` after removal. `Retained` findings carry a `Note`: the category was not selected, or removing the item would break the document.

## Categories by format

| Category | Word | Excel | PowerPoint |
| --- | --- | --- | --- |
| `DocumentProperties` | Author, last modified by, last printed, revision number, company, manager, template, total editing time, hyperlink base | Same | Same |
| `CustomProperties` | Removed; the part is deleted when empty | Same | Same |
| `Comments` | Comments, replies and their extension parts, with every range and reference mark | Notes and threaded comments with their replies | Classic and modern comments with replies |
| `TrackedChanges` | Reported per revision; accepted or rejected through the existing revision engine | Not applicable | Not applicable |
| `PersonalIdentities` | People part removed; authors on retained comments and revisions, including style and numbering revisions, anonymized | Threaded-comment persons anonymized with their user and provider IDs removed; note authors anonymized | Classic and modern comment authors anonymized |
| `HiddenContent` (opt-in) | Hidden text runs, including style-inherited hidden text; runs that carry field codes are skipped | Hidden and very hidden sheets, hidden rows and columns, hidden defined names and names pointing to hidden sheets. Hidden column ranges wider than 1,024 columns are retained, because they usually hide the unused part of a sheet | Hidden slides, hidden shapes and shapes placed entirely off the slide |
| `SpeakerNotes` | Not applicable | Not applicable | Notes slides removed |
| `Thumbnail` | Package thumbnail removed | Same | Same |
| `CustomXml` | Removed, except bibliography sources and parts bound to content controls | Removed | Removed |
| `ImageMetadata` | Exif, GPS, XMP and IPTC removed from JPEG, PNG, WebP, TIFF, BMP and GIF images without re-encoding pixels; ICC colour, density and orientation are kept. Other image formats are retained. While provenance is kept, a digital source type declaration stays and images with Content Credentials are retained | Same | Same |
| `Provenance` (opt-in) | C2PA manifests, external manifest references and IPTC digital source type declarations in JPEG, PNG, WebP, GIF, TIFF and SVG images, removed through the provenance engine. Package-level Content Credentials are always removed | Same | Same |
| `PrinterSettings` | Printer settings parts and the elements that reference them | Same | Same |
| `ExternalReferences` | Attached template removed. Other linked content, such as linked images and OLE links, is reported and retained | Cached values from external workbooks removed. External workbook paths and pivot cache source workbooks are reported and retained, because formulas and refreshes depend on them | Not applicable |
| `DocumentVariables` | Removed | Not applicable | Not applicable |
| `RevisionIdentifiers` | RSID attributes and the settings RSID table removed | Not applicable | Not applicable |

## Visio, OpenDocument and EPUB

`OdfDocument` covers all four OpenDocument kinds. `EpubPublication` is the editable EPUB model; `EpubDocument` is read-only and has no save. Categories a format does not have are never reported.

| Category | Visio | OpenDocument | EPUB |
| --- | --- | --- | --- |
| `DocumentProperties` | Author | Initial creator, last modified by, printed by, print date, editing cycles, editing duration, template and a generator string that carries build and platform | Creators, contributors, publishers and dates with their refinements. The EPUB 3 modification date is reported and kept, because the format requires it; saving an edited publication replaces it with the save time |
| `CustomProperties` | Not applicable | User-defined metadata | Named EPUB 2 `meta` values, except the cover declaration |
| `Comments` | Page comments | Annotations in text, cells, slides and pages, with their range ends | Not applicable |
| `TrackedChanges` | Not applicable | ODT insertions and deletions accepted or rejected through the document's own tracked-change methods; format changes can be accepted but not rejected. Spreadsheet tracked changes are reported and kept | Not applicable |
| `PersonalIdentities` | Comment authors and resolution identities anonymized | Authors and initials on kept annotations and tracked changes anonymized | Not applicable |
| `HiddenContent` (opt-in) | Hidden pages and hidden layers reported and kept: pages cannot be removed through the model, and shapes on a hidden layer can belong to visible layers | Hidden sections and hidden text or paragraph styles (ODT), hidden sheets (ODS), hidden slides and shapes (ODP) and shapes on hidden layers (ODG) removed. Conditional sections and hidden-text fields, and hidden rows and columns, are reported and kept. The last sheet or slide is kept | Not applicable |
| `SpeakerNotes` | Not applicable | Presentation notes removed | Not applicable |
| `Thumbnail` | Not saved | `Thumbnails/thumbnail.png` removed | Not applicable |
| `ImageMetadata` and `Provenance` | Embedded images on pages, shapes, connectors and masters, as for Word | Images in the package, as for Word. Package-level Content Credentials are always removed and the manifest entry goes with them | Manifest images, as for Word. Package-level Content Credentials are always removed with any manifest item that declares them |
| `PrinterSettings` | Not applicable | Printer name and driver setup in `settings.xml` removed | Not applicable |
| `ExternalReferences` | Hyperlinks to files and network paths reported and kept | Linked images, objects, sections and sheet sources reported and kept | Not applicable |

Saving a Visio drawing writes a new package from the drawing model. Package properties other than title and author, thumbnails and package-level Content Credentials of a loaded file are not written, so they are not reported.

## PDF

`PdfDocument` maps its content onto the same categories. Categories PDF does not have are never reported, and Word, Excel, PowerPoint, Visio, OpenDocument and EPUB never report `Attachments`, `ActiveContent` or `Bookmarks`.

| Category | PDF |
| --- | --- |
| `DocumentProperties` | User-authored Info fields and the catalog XMP packet. Producer, creation and modification dates, and trapping status are kept |
| `Comments` | Comment and markup annotations, except Link, Widget, file attachment and rich media annotations |
| `ActiveContent` | JavaScript, launch, submit-form, remote-navigation, import, movie, rendition and rich media actions; links and catalog base URIs whose scheme is not `http`, `https`, `mailto` or `tel`; and RichMedia, Movie, Sound, Screen and 3D annotations |
| `Attachments` (opt-in) | Embedded files, associated files and file attachment annotations |
| `Bookmarks` (opt-in) | The document outline |
| `HiddenContent` (opt-in) | Optional-content (layer) definitions and associations; the page content that belonged to a layer stays visible |
| `Provenance` | Content Credentials stored as associated files. They sign the whole file and cannot stay valid after the rewrite, so removal always removes them and says so in the finding's `Note`. Use `PdfProvenance.Remove` to change only provenance |
| `PersonalIdentities` | Not supported: comment authors are not anonymized. They appear in the comment finding's description and leave with the comment unless `Comments` is selected |

Removal rewrites the whole file, inspects the output again and throws when a selected item remains or when anything outside the selection changed. Signed PDFs are refused, because the rewrite would invalidate their signatures.

## Safety

Removal goes through the format owners: Word revisions use `AcceptRevisions`/`RejectRevisions` and comments use `WordComment.Remove`; Excel uses `ClearComment`, `RemoveThreadedComment`, `DeleteRows`, `DeleteColumns`, `RemoveWorksheet` and `RemoveNamedRange`; PowerPoint uses the comment, slide and shape `Remove` methods. Package-level parts are removed with their relationships and the elements that reference them. Tests reopen every scrubbed file, validate it against the Microsoft 365 Open XML schema and inspect it again.

Deleting hidden rows, columns or sheets shifts or breaks formulas that refer to them, exactly as deleting them in Excel does. Removing hidden content is therefore never part of `Default`. Signed packages are not re-signed; save a scrubbed document before signing it.
