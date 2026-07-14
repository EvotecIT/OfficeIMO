# OfficeIMO.Markup.PowerPoint - Markup to PowerPoint export

`OfficeIMO.Markup.PowerPoint` exports the semantic `OfficeIMO.Markup` presentation model to editable PowerPoint `.pptx` files through `OfficeIMO.PowerPoint`.

This project is built from the OfficeIMO source tree and is not published as a standalone NuGet package.

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

result.Document.SaveAsPowerPoint("quarterly-review.pptx", new MarkupToPowerPointOptions {
});
```

## What it exports

- Slides, real PowerPoint sections, titles, text, lists, and tables.
- Images and local background images resolved relative to the markup file when an input path is supplied.
- Native linear-gradient backgrounds, overlays, speaker notes, and transition metadata.
- Native charts from inline CSV chart data.
- Optional Mermaid-to-image export when Mermaid CLI is available.

## Mermaid rendering

Set `MermaidRendererPath` in `MarkupToPowerPointOptions`, pass `--mermaid-renderer <path-to-mmdc>` through `OfficeIMO.Markup.Cli`, or set `OFFICEIMO_MARKUP_MERMAID_CLI`.

## Boundaries

- Markup parsing and validation stay in `OfficeIMO.Markup`.
- Presentation creation stays in `OfficeIMO.PowerPoint`.
- This package maps semantic presentation nodes into editable PowerPoint output.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None by default. Mermaid image export can use a caller-installed Mermaid CLI.
- **OfficeIMO:** `OfficeIMO.Markup`, `OfficeIMO.PowerPoint`, and `OfficeIMO.Drawing`; the exporter maps semantic nodes to editable slide content.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
