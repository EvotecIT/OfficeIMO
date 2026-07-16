# OfficeIMO.Reader.OneNote

`OfficeIMO.Reader.OneNote` projects native, offline `OfficeIMO.OneNote` content into the shared `OfficeIMO.Reader` contracts. It does not use Microsoft Graph, GraphEssentialsX, COM automation, or require OneNote to be installed.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.OneNote;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddOneNoteHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("Notes.one");
```

The adapter emits page-aware text and Markdown chunks, structured tables, links, and image/embedded-file asset metadata. Native file parsing remains owned by `OfficeIMO.OneNote`.

Current pages are the default ingestion surface. Conflict copies and version-history snapshots remain counted in metadata but enter chunks, page hierarchy, links, and assets only when requested:

```csharp
using OfficeIMO.OneNote;

var oneNoteOptions = new ReaderOneNoteOptions {
    IncludeConflictPages = true,
    IncludeVersionHistory = true,
    IncludeAssetPayloads = true,
    NotebookOptions = new OneNoteNotebookReaderOptions {
        IncludeRecycleBin = true
    }
};

OfficeDocumentReader readerWithHistory = new OfficeDocumentReaderBuilder()
    .AddOneNoteHandler(oneNoteOptions)
    .Build();
```

Image metadata is emitted even when a native image payload cannot be resolved. Resolved payload identifiers remain stable because unresolved images still occupy their native position. Picture dimensions are reported in pixels after converting MS-ONE half-inch units at 96 DPI. Native RichEdit separators and invalid control/noncharacter code points are normalized in projected text without mutating the source model.

Before building chunks, tables, assets, links, or metadata, the adapter applies the shared bounded projection validation to the complete model graph. This includes conflict and version relationships because Reader counts them in metadata even when their content is not selected. Caller-created cycles, shared instances, or excessive section-group, related-page, and content depth therefore fail predictably instead of causing recursive traversal or repeated expansion.

Current target frameworks are `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows. The package is MIT licensed.
