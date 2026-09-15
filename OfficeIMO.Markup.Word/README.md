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
- **OfficeIMO:** `OfficeIMO.Markup`, `OfficeIMO.Word`, and `OfficeIMO.Core`; the exporter maps semantic nodes to editable Word content.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Markup.Word` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
