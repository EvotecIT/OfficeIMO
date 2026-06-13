# OfficeIMO.Markup.Word - Markup to Word export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Markup.Word)](https://www.nuget.org/packages/OfficeIMO.Markup.Word)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markup.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Markup.Word)

`OfficeIMO.Markup.Word` exports the semantic `OfficeIMO.Markup` document model to editable Word `.docx` files through `OfficeIMO.Word`.

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
