# Hidden data and personal information support matrix

`WordDocument`, `ExcelDocument` and `PowerPointDocument` find and remove the metadata and hidden data that Office's Document Inspector covers. `InspectHiddenData()` changes nothing and returns an `OfficeHiddenDataReport`. `RemoveHiddenData(OfficeHiddenDataOptions)` removes the selected categories in memory and returns the same report shape, where each finding records what happened to it. Save the document to keep the result.

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

`OfficeHiddenDataKind.Default` selects every category except tracked changes and hidden content, which change what readers see, and provenance, which discloses how content was made. Tracked changes are reported and kept unless `TrackedChanges` is `AcceptAll` or `RejectAll`. Authors on comments and revisions that stay in the document are replaced with `AnonymousAuthor` (default `Author`) rather than deleted, because Office requires an author name on them.

## Provenance

The report also answers what provenance leaves with the file. Content Credentials (C2PA manifests and references to them) and IPTC digital source type declarations, such as an image marked as AI-generated, appear as `Provenance` findings whose `Provenance` property carries the same evidence that `InspectProvenance` returns.

Provenance is kept unless you add `OfficeHiddenDataKind.Provenance`. While it is kept, `ImageMetadata` cleanup reduces an image's XMP to its digital source type declaration instead of deleting it, and leaves images that carry Content Credentials untouched, because changing their bytes would invalidate the credential; those images are reported as `Retained`. Selecting `Provenance` removes the carriers through the same engine as `RemoveProvenance`, and then image metadata is stripped as usual.

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

## Safety

Removal goes through the format owners: Word revisions use `AcceptRevisions`/`RejectRevisions` and comments use `WordComment.Remove`; Excel uses `ClearComment`, `RemoveThreadedComment`, `DeleteRows`, `DeleteColumns`, `RemoveWorksheet` and `RemoveNamedRange`; PowerPoint uses the comment, slide and shape `Remove` methods. Package-level parts are removed with their relationships and the elements that reference them. Tests reopen every scrubbed file, validate it against the Microsoft 365 Open XML schema and inspect it again.

Deleting hidden rows, columns or sheets shifts or breaks formulas that refer to them, exactly as deleting them in Excel does. Removing hidden content is therefore never part of `Default`. Signed packages are not re-signed; save a scrubbed document before signing it.
