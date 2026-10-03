# OfficeIMO.Reader.Latex

This modular adapter adds `.tex` ingestion backed by `OfficeIMO.Latex`. It extracts the bounded OfficeIMO profile and diagnoses preserved or simplified content; it does not compile TeX or load packages.

PreserveOnly mode carries the complete visible source, including preamble declarations, through block and whole-document chunks. Description items without a visible term retain their body without an added colon; label anchors remain available when chunks split.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Latex;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddLatexHandler()
    .Build();
IReadOnlyList<ReaderChunk> chunks = reader.Read("article.tex").ToList();
```

Chunks retain source locations, heading hierarchy, block kind, Markdown projection, and parser/conversion warnings. Article, report, and book documents get ordered typed chunks for headings, paragraphs, lists (including description lists), figures with captions, tables, theorems, and math. Whole-document mode projects the same supported blocks instead of collapsing the file to paragraphs. Plain TeX or another document class receives an unrecognized-profile warning and a visible source fallback instead of empty output.

List items, definitions and quotations retain multiple paragraphs and nested block children. Footnotes include a typed reference and a separate definition body; block mode identifies the latter as `SourceBlockKind = "footnote"` and keeps its source heading path. Explicit TeX marks and nested-note insertions produce fidelity warnings. Split chunks keep the note text while reporting flattened layout.

Markdown-only content, such as an anchor without visible text, remains available in chunks. Metadata and preamble warnings reach the first emitted chunk once, including metadata-only and whole-document output. A figure's shared caption appears once in extracted text alongside its image paths.

When `MaxChars` splits content, each chunk retains literal prose, protected code or math source, and complete label anchors. Split Markdown flattens layout and formatting and reports that simplification. `MaxChars` bounds extracted text on a best-effort basis; Markdown escaping and code fences add carrier overhead, and an indivisible anchor or Unicode scalar may exceed a very small limit. Literal HTML in code or prose remains text; only projected LaTeX labels become anchors.

The handler enforces the smaller of Reader and native input limits for both seekable and non-seekable streams. Native loading reads the caller's stream directly and preserves its seekable position without an extra adapter snapshot. Parsing and writing remain owned by `OfficeIMO.Latex`.

## Dependency footprint

- **External:** None; no TeX runtime or compiler.
- **OfficeIMO:** `OfficeIMO.Reader.Core`, `OfficeIMO.Latex`, and `OfficeIMO.Latex.Markdown`; parsing stays in the native LaTeX package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
