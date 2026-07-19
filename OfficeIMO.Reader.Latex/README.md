# OfficeIMO.Reader.Latex

This modular adapter adds `.tex` ingestion backed by `OfficeIMO.Latex`. It extracts the bounded OfficeIMO profile and diagnoses preserved or simplified content; it does not compile TeX or load packages.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Latex;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddLatexHandler()
    .Build();
IReadOnlyList<ReaderChunk> chunks = reader.Read("article.tex").ToList();
```

Chunks retain source locations, heading hierarchy, block kind, Markdown projection, and parser/conversion warnings. Article, report, and book documents get ordered typed chunks for headings, paragraphs, lists (including description lists), figures with captions, tables, theorems, and math. Whole-document mode projects the same supported blocks instead of collapsing the file to paragraphs. Plain TeX or another document class receives an unrecognized-profile warning and a visible source fallback instead of empty output.

The handler enforces Reader input limits for both seekable and non-seekable streams. Parsing and writing remain owned by `OfficeIMO.Latex`.

## Dependency footprint

- **External:** None; no TeX runtime or compiler.
- **OfficeIMO:** `OfficeIMO.Reader.Core`, `OfficeIMO.Latex`, and `OfficeIMO.Latex.Markdown`; parsing stays in the native LaTeX package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
