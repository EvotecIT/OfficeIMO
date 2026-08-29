# OfficeIMO.Reader.DocBook

`OfficeIMO.Reader.DocBook` registers deterministic `.dbk` and `.docbook` ingestion backed by the source-preserving, bounded `OfficeIMO.DocBook` package. Generic `.xml` remains routed to `OfficeIMO.Reader.Xml`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.DocBook;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddDocBookHandler()
    .Build();

IReadOnlyList<ReaderChunk> chunks = reader.Read("guide.docbook").ToList();
```

The adapter emits common-structure chunks with stable IDs, section paths, source kinds, simple Markdown projections, and optional bounded-profile warnings. Parsing, schema-profile identification, limits, validation, editing, and extension preservation remain owned by `OfficeIMO.DocBook`.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.DocBook`.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
