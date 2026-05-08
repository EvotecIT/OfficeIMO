# OfficeIMO.Markup.Word

`OfficeIMO.Markup.Word` exports the semantic `OfficeIMO.Markup` document model to editable Word `.docx` files through `OfficeIMO.Word`.

Use it when a `.omd` or Markdown-inspired authoring file has `profile: document` and should become a real Word document rather than generated C# or PowerShell starter code.

## Install

```powershell
dotnet add package OfficeIMO.Markup.Word
```

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

new OfficeMarkupWordExporter().Export(result.Document, new OfficeMarkupWordExportOptions {
    OutputPath = "status-brief.docx"
});
```

## What exports today

- Headings, paragraphs, lists, and pipe tables
- Images resolved relative to the markup file when an input path is supplied
- Page breaks, sections, headers, footers, and table of contents directives
- Inline chart data mapped to native Word chart output

## Related packages

- `OfficeIMO.Markup`: parser, semantic AST, validation, and emitters
- `OfficeIMO.Markup.Cli`: command-line parse, validate, emit, and export workflow
- `OfficeIMO.Word`: Word document object model used by this exporter

## Targets

- `netstandard2.0`, `net8.0`, `net10.0`
- `net472` when building on Windows

License: MIT
