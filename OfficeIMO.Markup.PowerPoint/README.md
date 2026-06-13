# OfficeIMO.Markup.PowerPoint - Markup to PowerPoint export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Markup.PowerPoint)](https://www.nuget.org/packages/OfficeIMO.Markup.PowerPoint)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markup.PowerPoint?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Markup.PowerPoint)

`OfficeIMO.Markup.PowerPoint` exports the semantic `OfficeIMO.Markup` presentation model to editable PowerPoint `.pptx` files through `OfficeIMO.PowerPoint`.

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

## What it exports

- Slides, real PowerPoint sections, titles, text, lists, and tables.
- Images and local background images resolved relative to the markup file when an input path is supplied.
- Native linear-gradient backgrounds, overlays, speaker notes, and transition metadata.
- Native charts from inline CSV chart data.
- Optional Mermaid-to-image export when Mermaid CLI is available.

## Mermaid rendering

Set `MermaidRendererPath` in `OfficeMarkupPowerPointExportOptions`, pass `--mermaid-renderer <path-to-mmdc>` through `OfficeIMO.Markup.Cli`, or set `OFFICEIMO_MARKUP_MERMAID_CLI`.

## Boundaries

- Markup parsing and validation stay in `OfficeIMO.Markup`.
- Presentation creation stays in `OfficeIMO.PowerPoint`.
- This package maps semantic presentation nodes into editable PowerPoint output.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
