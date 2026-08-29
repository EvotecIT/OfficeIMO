# OfficeIMO.Reader.Opml

`OfficeIMO.Reader.Opml` registers deterministic `.opml` ingestion backed by the lossless, bounded `OfficeIMO.Opml` package.

## Install

```shell
dotnet add package OfficeIMO.Reader.Opml
```

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Opml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddOpmlHandler()
    .Build();

IReadOnlyList<ReaderChunk> chunks = reader.Read("subscriptions.opml").ToList();

OfficeDocumentReadResult document = reader.ReadDocument("subscriptions.opml");
```

`Read` emits one or more bounded chunks per outline with stable IDs, nesting-aware heading paths, and optional OPML validation warnings. `ReadDocument` also publishes head metadata, subscription and outline links, and each native or conversion diagnostic once at document scope. Parsing, limits, validation, and the editable model remain owned by `OfficeIMO.Opml`.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.Opml`.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
