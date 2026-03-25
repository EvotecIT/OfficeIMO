---
title: Visio Diagrams
description: Overview of the OfficeIMO.Visio package for generating VSDX diagrams with pages, shapes, and connectors.
order: 52
---

# Visio Diagrams

`OfficeIMO.Visio` focuses on generating `.vsdx` diagrams from pure .NET code. It is a good fit when your application already knows the topology, flow, or system shape it wants to draw and you need a repeatable way to emit diagrams without manual Visio authoring.

## Good use cases

- Infrastructure and architecture diagrams generated from configuration or inventory data.
- Network maps, dependency graphs, and deployment flows.
- Org charts or process diagrams produced from business system exports.
- Report pipelines that need a diagram artifact alongside Word, Excel, or PowerPoint output.

## Building blocks

- **Diagrams and pages** to organize one or more related views.
- **Shapes** with text, size, fill color, and placement.
- **Connectors** for relationships and directional flows between shapes.
- **Measurement helpers** so positioning can stay readable in code.

## Quick start

```csharp
using OfficeIMO.Visio;

using var diagram = VisioDiagram.Create("topology.vsdx");
var page = diagram.AddPage("System Overview");

var frontend = page.AddShape("Frontend", 2.0, 8.0);
var api = page.AddShape("API", 5.0, 8.0);
var database = page.AddShape("Database", 8.0, 8.0);

page.AddConnector(frontend, api, "HTTPS");
page.AddConnector(api, database, "SQL");

diagram.Save();
```

## Recommended workflow

1. Define the nodes you want to visualize from your source data.
2. Map each node to a stable shape position and visual style.
3. Add connectors that represent traffic, dependency, or process flow.
4. Save the `.vsdx` output as a build artifact or generated report attachment.

## Related packages

- [OfficeIMO.Word](/products/word/) for report documents that embed or reference generated diagrams.
- [OfficeIMO.PowerPoint](/products/powerpoint/) for slide decks built from the same topology data.
- [OfficeIMO.Reader](/products/reader/) when your workflow also needs ingestion and extraction.
