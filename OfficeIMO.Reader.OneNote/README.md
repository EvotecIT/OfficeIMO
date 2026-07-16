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

Current target frameworks are `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows. The package is MIT licensed.
