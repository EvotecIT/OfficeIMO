---
title: "Create and edit Visio diagrams in .NET"
description: "Create, edit, inspect, validate, and export VSDX diagrams with builders, pages, shapes, connectors, stencils, and graph layouts."
meta.seo_title: "Create, edit, and export VSDX Visio diagrams in .NET"
order: 52
---

`OfficeIMO.Visio` creates, edits, inspects, validates, and exports `.vsdx` diagrams without COM automation and without Microsoft Visio installed. Use a diagram builder for common topology and process models, or work directly with pages, shapes, connectors, layers, Shape Data, containers, comments, and metadata.

## Choose your workflow

| You need to | Start here |
|---|---|
| Generate flowcharts, architectures, networks, org charts, sequences, timelines, or generic graphs | [Diagram builders](/docs/visio/diagram-builders/) |
| Load a VSDX, select and update shapes, maintain metadata, or validate output | [Editing and validation](/docs/visio/editing-and-validation/) |
| Use built-in, installed, or external VSSX/VSTX shapes | [Stencils and catalogs](/docs/visio/stencils-and-catalogs/) |
| Produce SVG, PNG, JPEG, TIFF, or WebP previews without Visio | [Image export](/docs/visio/image-export/) |

## What the package handles

- VSDX document and page creation, loading, editing, saving, and validation
- shapes, text, styles, connection points, connectors, routing hints, labels, hyperlinks, and Shape Data
- layers, containers, comments, metadata, and topology queries
- fluent builders for common business and technical diagram types
- generated catalogs, installed Visio stencils, external VSSX/VSTX packs, and stencil migration planning
- headless SVG, PNG, JPEG, TIFF, and lossless WebP export

## Quick start

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("topology.vsdx")
    .ArchitectureDiagram("System Overview", diagram => diagram
        .Title()
        .Legend()
        .Theme(VisioStyleTheme.Technical())
        .Actor("users", "Users", 0, 1)
        .Gateway("gateway", "Gateway", 1, 1)
        .Service("api", "API", 2, 1)
        .Database("database", "Database", 3, 1)
        .DataFlow("users", "gateway", "HTTPS")
        .ControlFlow("gateway", "api", "route")
        .DataFlow("api", "database", "SQL"))
    .Save();
```

Run the [architecture example](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Visio/ArchitectureDiagramBuilder.cs) to inspect a generated VSDX and validation result. The [Visio API reference](/api/visio/) contains the complete lower-level model.
