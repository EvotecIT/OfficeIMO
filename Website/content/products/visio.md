---
title: "OfficeIMO.Visio"
description: "Create Visio diagrams with shapes, connectors, and pages. No Visio installation required."
layout: product
product_color: "#ea580c"
install: "dotnet add package OfficeIMO.Visio"
nuget: "OfficeIMO.Visio"
docs_url: "/docs/visio/"
api_url: ""
---

## Why OfficeIMO.Visio?

OfficeIMO.Visio lets you generate and modify `.vsdx` diagrams from pure .NET code. Build network topologies, org charts, flowcharts, and architecture diagrams without requiring Visio on the machine. The fluent builder API keeps diagram construction concise and readable.

## Features

- **Create, load & save .vsdx files** -- full round-trip support for the Visio Open XML format
- **Pages & shapes** -- add pages, drop shapes from stencils, and set position, size, text, and fill color
- **Connectors with auto-glue** -- connect shapes with dynamic connectors that route automatically
- **Fluent builder API** -- chain calls to construct diagrams in a single expression
- **Multiple measurement units** -- work in inches, millimeters, centimeters, or points
- **Connection points** -- define and target specific connection points on shapes for precise routing

## Diagram types you can automate

| Diagram type | Typical source data | Outcome |
|--------------|---------------------|---------|
| Infrastructure topology | Inventory exports, deployment metadata, cloud resources | Repeatable architecture diagrams that stay in sync with real systems |
| Process and approval flows | Workflow definitions, queue stages, or policy steps | Clear step-by-step diagrams for operations and compliance teams |
| Org and service maps | HR systems, ownership registries, service catalogs | Diagrams that show teams, responsibilities, and platform boundaries |
| Release and runbook visuals | CI pipelines, deployment targets, or support procedures | Supporting diagrams that travel with Word reports and PowerPoint decks |

## Quick start

```csharp
using OfficeIMO.Visio;

using var diagram = VisioDiagram.Create("Architecture.vsdx");
var page = diagram.AddPage("System Overview");

// Add shapes
var webServer = page.AddShape("Web Server", 2.0, 8.0);
webServer.FillColor = "#2563eb";
webServer.TextColor = "#ffffff";

var appServer = page.AddShape("App Server", 5.0, 8.0);
appServer.FillColor = "#059669";
appServer.TextColor = "#ffffff";

var database = page.AddShape("Database", 8.0, 8.0);
database.FillColor = "#7c3aed";
database.TextColor = "#ffffff";

var cache = page.AddShape("Cache", 5.0, 5.5);
cache.FillColor = "#dc2626";
cache.TextColor = "#ffffff";

// Connect shapes
page.AddConnector(webServer, appServer, "HTTPS");
page.AddConnector(appServer, database, "TCP/IP");
page.AddConnector(appServer, cache, "Redis");

diagram.Save();
```

## Repeatable modeling flow

1. Start from the system, process, or ownership data you already have instead of hand-placing every shape.
2. Map each entity to a stable shape type, label, and color so regenerated diagrams stay easy to compare.
3. Use connectors to represent traffic, dependency, or approval flow, then reserve shape text for the labels people need to scan quickly.
4. Generate the `.vsdx` file as part of a report build, architecture export, or operational handoff package.
5. Pair the diagram with OfficeIMO.Word or OfficeIMO.PowerPoint when the same workflow also needs narrative or presentation output.

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.Visio runs on Windows, Linux, and macOS. Generated files open in Microsoft Visio and other VSDX-compatible viewers.

## Related guides

| Guide | Description |
|-------|-------------|
| [Visio documentation](/docs/visio/) | Review the shape, page, and connector model before you build diagrams. |
| [Getting Started](/docs/getting-started/) | Set up the package and validate your first generated diagram. |
| [OfficeIMO.Word](/products/word/) | Pair diagrams with generated reports, architecture notes, or implementation docs. |
| [OfficeIMO.PowerPoint](/products/powerpoint/) | Turn the same topology or process data into stakeholder-ready slide decks. |
