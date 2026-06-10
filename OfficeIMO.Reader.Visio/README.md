# OfficeIMO.Reader.Visio

Thin Visio adapter for `OfficeIMO.Reader`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;

DocumentReaderVisioRegistrationExtensions.RegisterVisioHandler();

IReadOnlyList<ReaderChunk> chunks = DocumentReader
    .Read("diagram.vsdx")
    .ToList();
```

The adapter uses `OfficeIMO.Visio` inspection snapshots and emits page-aware
chunks for `.vsdx`, `.vsdm`, `.vstx`, and `.vstm` files. Shape Data is exposed
as `ReaderTable` rows, and `ReadVisioDocument(...)` maps pages, shapes,
connectors, hyperlinks, and optional preview asset metadata into the shared
OfficeIMO document read result envelope.
