# OfficeIMO.Reader.Word

Word ingestion for `OfficeIMO.Reader.Core` with DOCX, DOCM, and legacy DOC support.

## Project an open document

Import `OfficeIMO.Reader.Word` and call `document.ReadDocument(sourceName,
options, formatOptions, cancellationToken)` on a caller-owned `WordDocument`.
The method uses the same Reader projection as path and stream reads, snapshots
both option sets and enforces the projection's output budgets. `Full` is the
default. The model remains open for subsequent `InspectEmbeddedPayloads` reads.

The source name supplies citations without reopening a file or inventing its
byte hash, length or timestamp. Input and package admission settings cannot
retroactively constrain the earlier native load; configure them when loading
the model. Lean projection still requires a modern XML source model.

## Text and metadata projection

Use `ReaderOptions.Projection = ReaderProjection.TextAndMetadata` for DOCX, DOCM,
DOTX and DOTM inspection when image bytes, computed page layout and Markdown are
not requested. The adapter retains semantic text, tables, links and requested
native inspection metadata. It omits image payload copies before snapshot and
asset construction, skips Markdown rendering, and leaves computed pages absent.
This projection does not infer physical pages for flow text.

With `Inspection.IncludePackageContent`, lean reads inventory inert payload
descriptors through a bounded native session. Sizes unavailable without reading
remain explicitly unknown. VBA part/signature-relationship metadata is retained,
while module decoding is marked `module-payload-unrequested`.

Binary DOC/DOT and the separate legacy word-processing handlers require the
ordinary `Full` projection. They refuse `TextAndMetadata` rather than silently
render graphics or construct a full import. Native routing also rejects a binary
source mislabeled with a modern extension. Native parser limits remain separate
from Reader's aggregate result budgets.

`OpenXmlMaxPartUncompressedBytes` and `OpenXmlMaxTotalUncompressedBytes` apply
native declared part/aggregate expansion admission before model loading, even
when `Inspection` is null. Configured ceilings produce terminal
`ReaderResourceLimitException` with the native failure retained; they are not
retried as password failures. These are package admission bounds, not a hard
limit on process memory or every decoder's actual work.


Rich `list-item` blocks carry the resolved Word marker in `OfficeDocumentBlock.Marker`,
including authored starts, Roman or lettered formats, continuation, and explicit restarts.
Portable bullet symbols are retained, and invisible list markers are empty.
The values come from `WordParagraphSnapshot.ListMarker`; inspection snapshots also expose
the numeric counter of an ordered item through `WordParagraphSnapshot.ListIndex`.
Inspection and ingestion reject list markers longer than 4,096 characters or more than
4,194,304 generated marker characters in a document with an `InvalidDataException`.

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .Build();
```

To add best-effort page locations for search and citations:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler(new ReaderWordOptions {
        IncludePageLocations = true
    })
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("policy.docx");
OfficeDocumentSearchResult matches = document.Search("retention period");

foreach (int page in matches.PageNumbers) {
    Console.WriteLine($"Found on page {page} of {matches.TotalPageCount}");
}
```

Word does not store stable physical pages. Enabling `IncludePageLocations` uses the same OfficeIMO.Word page layout
as PDF and image export and reports `OfficeDocumentPageProvenance.Computed`. `PageLocationOptions.LayoutOptions`
configures that layout. The normalized `Blocks` and `Pages` contain visible rendered text lines with page-specific
identities and baseline extents in PDF user-space points. Logical text remains available through `Chunks`; tables
and links remain at document scope because rendered lines do not identify their authored source blocks. Page Markdown
therefore follows rendered text rather than authored paragraph or list structure. Results can differ from Microsoft Word
when fonts, metrics, or unsupported layout features differ, so the option is disabled by default.

Selected WordPerfect, WordStar, Ami Pro, Lotus Word Pro, Works/Write, and Word for DOS sources are included in `OfficeIMO.Word` but remain opt-in at the Reader registration boundary:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordAndLegacyHandlers(new LegacyWordImportOptions {
        Limits = new OfficeLegacyImportLimits { MaxInputBytes = 64 * 1024 * 1024 }
    })
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("archive.wpd");
```

The combined registration applies the same immutable legacy options to both the family-specific extensions and content-routed Word for DOS `.doc`; compound-binary `.doc` remains on the normal Word path. `AddLegacyWordHandler(...)` remains available when only the unambiguous legacy extensions are wanted and deliberately does not claim `.doc`.

Legacy warnings include the detected profile, structured-versus-salvage quality, and feature-level losses. The handler never executes source macros or embedded code.

## Inspect comments and tracked changes

Use a rich read to retain native review and package evidence:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

var reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
OfficeDocumentReadResult document = reader.ReadDocument("review.docx", new ReaderOptions {
    Inspection = new ReaderInspectionOptions()
});
foreach (OfficeDocumentMetadataEntry entry in document.Metadata) {
    if (entry.Category == OfficeDocumentInspectionCategories.Comment ||
        entry.Category == OfficeDocumentInspectionCategories.Revision)
        Console.WriteLine($"{entry.Category}: {entry.Value} ({entry.SourceObjectId})");
}
```

Comment records retain available authors, dates, reply/resolution state, target text
and story identifiers. Revision records retain affected text, source part and story;
inspection does not accept or reject changes. Unsupported review metadata remains
explicit in coverage and diagnostics. Visibility observations use the content-safety
owner; requesting visibility while disabling content safety does not infer hidden
state from the extracted text.

Package inspection reads the loader's source bytes and retains inert payload/VBA
descriptors. Loaded-content safety inspection does not reopen or serialize the
document. [Reader Core](../OfficeIMO.Reader.Core/README.md#inspect-review-and-saved-content)
defines the shared record/character ceilings and terminal resource failures.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
- OfficeIMO dependencies: `OfficeIMO.Reader.Core`, `OfficeIMO.Word`, and `OfficeIMO.Core`.
- License: MIT.
