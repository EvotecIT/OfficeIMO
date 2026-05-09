# OfficeIMO.Markup.PowerPoint

`OfficeIMO.Markup.PowerPoint` exports the semantic `OfficeIMO.Markup` presentation model to editable PowerPoint `.pptx` files through `OfficeIMO.PowerPoint`.

Use it when a `.omd` or Markdown-inspired authoring file has `profile: presentation` and should become a native deck with slides, sections, text, tables, images, charts, backgrounds, speaker notes, and transitions.

## Install

```powershell
dotnet add package OfficeIMO.Markup.PowerPoint
```

## Quick start

```csharp
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;

var result = OfficeMarkupParser.Parse("""
---
profile: presentation
title: Quarterly Review
---

# Quarterly Review

@slide {
  layout: title-and-content
  transition: fade
}

- Revenue grew
- Churn improved

::notes
Open with the top-line result.
""");

new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
    OutputPath = "quarterly-review.pptx"
});
```

## What exports today

- Slides, real PowerPoint sections, titles, text, lists, and tables
- Images and local background images resolved relative to the markup file when an input path is supplied
- Native linear-gradient backgrounds, overlays, speaker notes, and transition metadata
- Native charts from inline CSV chart data
- Optional Mermaid-to-image export when Mermaid CLI is available

## Mermaid rendering

Set `MermaidRendererPath` in `OfficeMarkupPowerPointExportOptions`, pass `--mermaid-renderer <path-to-mmdc>` through `OfficeIMO.Markup.Cli`, or set `OFFICEIMO_MARKUP_MERMAID_CLI`.

## Related packages

- `OfficeIMO.Markup`: parser, semantic AST, validation, and emitters
- `OfficeIMO.Markup.Cli`: command-line parse, validate, emit, and export workflow
- `OfficeIMO.PowerPoint`: PowerPoint object model used by this exporter

## Targets

- `netstandard2.0`, `net8.0`, `net10.0`
- `net472` when building on Windows

License: MIT
