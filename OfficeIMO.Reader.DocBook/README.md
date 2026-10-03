# OfficeIMO.Reader.DocBook

`OfficeIMO.Reader.DocBook` registers deterministic `.dbk` and `.docbook` ingestion backed by the source-preserving, bounded `OfficeIMO.DocBook` package. Generic `.xml` remains routed to `OfficeIMO.Reader.Xml`.

## Install

```shell
dotnet add package OfficeIMO.Reader.DocBook
```

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.DocBook;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddDocBookHandler()
    .Build();

IReadOnlyList<ReaderChunk> chunks = reader.Read("guide.docbook").ToList();

OfficeDocumentReadResult document = reader.ReadDocument("guide.docbook");
```

`Read` emits common-structure chunks with stable IDs, section paths, source kinds, admonition context, list-aware Markdown indentation, inline link and cross-reference targets, content-safe fences for program listings and screens, and optional bounded-profile warnings. `ReadDocument` also publishes source metadata, authors, links, bounded CALS tables, and each native or conversion diagnostic once at document scope. `ReaderDocBookOptions.ReadOptions` controls native parsing bounds and `ConversionOptions` controls shared-model text, table, and diagnostic budgets. Parsing, schema-profile identification, validation, editing, and extension preservation remain owned by `OfficeIMO.DocBook`.

Ordinary source text is escaped as literal Markdown, including text resembling emphasis, links, headings or lists. Native emphasis and literal spans retain their formatting; emphasis with significant boundary whitespace or split around links uses inline HTML. Program listings and screens retain their content-safe fences. Flattened CALS cells separate distinct block paragraphs with newlines while keeping adjacent inline text together. Escapes and generated markup count toward Reader chunk limits.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.DocBook`.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
