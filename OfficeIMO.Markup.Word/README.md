# OfficeIMO.Markup.Word - Markup to Word export

`OfficeIMO.Markup.Word` exports the semantic `OfficeIMO.Markup` document model to editable Word `.docx` files through `OfficeIMO.Word`.

This project is built from the OfficeIMO source tree and is not published as a standalone NuGet package.

## Quick start

```csharp
using OfficeIMO.Markup;
using OfficeIMO.Markup.Word;

var result = OfficeMarkupParser.Parse("""
---
profile: document
title: Status Brief
---

# Status Brief

This document was authored as OfficeIMO Markup.

::pagebreak

## Appendix
Generated as an editable Word document.
""");

result.Document.SaveAsWord("status-brief.docx", new MarkupToWordOptions {
});
```

## What it exports

- Headings, paragraphs, lists, and pipe tables.
- Images resolved relative to the markup file when an input path is supplied.
- Page breaks, sections, headers, footers, and table-of-contents directives.
- Inline chart data mapped to native Word chart output.

## Boundaries

- Markup parsing and validation stay in `OfficeIMO.Markup`.
- Word document creation stays in `OfficeIMO.Word`.
- This package maps semantic document nodes into editable Word output.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Markup`, `OfficeIMO.Word`, and `OfficeIMO.Drawing`; the exporter maps semantic nodes to editable Word content.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
