# OfficeIMO.Reader.AsciiDoc

`OfficeIMO.Reader.AsciiDoc` provides a modular `.adoc`, `.asciidoc`, and `.asc` handler backed by the dependency-free, lossless `OfficeIMO.AsciiDoc` engine.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.AsciiDoc;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAsciiDocHandler()
    .Build();
IReadOnlyList<ReaderChunk> chunks = reader.Read("guide.adoc").ToList();
```

The handler emits deterministic block-aware chunks with source lines, heading paths, typed Markdown projections, structured table content, compound-list ownership, and parser/conversion warnings. Whole-document mode includes attached list content once, without duplicating it as a top-level block. Parsing and writing remain owned by `OfficeIMO.AsciiDoc`; this package only adapts the native model to Reader contracts.
