---
title: Visio Diagrams
description: Generate VSDX diagrams with builders, pages, shapes, connectors, stencils, and graph layouts.
order: 52
---

`OfficeIMO.Visio` focuses on generating `.vsdx` diagrams from pure .NET code. It is a good fit when your application already knows the topology, flow, or system shape it wants to draw and you need a repeatable way to emit diagrams without manual Visio authoring.

## Good use cases

- Infrastructure and architecture diagrams generated from configuration or inventory data.
- Network maps, dependency graphs, and deployment flows.
- Org charts, timelines, sequences, swimlanes, or process diagrams produced from business system exports.
- Native or external stencil-pack diagrams generated from installed Visio stencils or `.vssx` repositories.
- Report pipelines that need a diagram artifact alongside Word, Excel, or PowerPoint output.

## Building blocks

- **Documents and pages** to organize one or more related views.
- **Shapes, masters, stencils, and catalogs** for generated, native, and external-pack visuals.
- **Connectors** for relationships, directional flows, labels, Shape Data, and hyperlinks.
- **Diagram builders** for flowcharts, architecture, networks, dependencies, swimlanes, org charts, timelines, sequences, and generic graphs.
- **Quality and polish helpers** for page fitting, text sizing, connector labels, and visual validation.

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

## Recommended workflow

1. Define the nodes and relationships you want to visualize from your source data.
2. Choose a domain builder, such as architecture, network, flowchart, dependency, swimlane, timeline, or graph.
3. Use generated catalogs, installed Visio stencils, or external `.vssx` packs when real domain symbols matter.
4. Attach Shape Data, hyperlinks, and stable IDs so regenerated diagrams remain searchable and diff-friendly.
5. Save the `.vsdx` output as a build artifact or generated report attachment.

## Related packages

- [OfficeIMO.Visio API reference](/api/visio/) for diagram, shape, connector, and fluent builder types.
- [OfficeIMO.Word](/products/word/) for report documents that embed or reference generated diagrams.
- [OfficeIMO.PowerPoint](/products/powerpoint/) for slide decks built from the same topology data.
- [OfficeIMO.Reader](/products/reader/) when your workflow also needs ingestion and extraction.
