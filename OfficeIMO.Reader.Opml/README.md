# OfficeIMO.Reader.Opml

`OfficeIMO.Reader.Opml` registers deterministic `.opml` ingestion backed by the lossless, bounded `OfficeIMO.Opml` package.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Opml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddOpmlHandler()
    .Build();

IReadOnlyList<ReaderChunk> chunks = reader.Read("subscriptions.opml").ToList();
```

One or more bounded chunks are emitted per outline with stable IDs, nesting-aware heading paths, and optional OPML validation warnings. Parsing, limits, validation, and the editable model remain owned by `OfficeIMO.Opml`.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.Opml`.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
