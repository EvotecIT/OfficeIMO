---
title: "OfficeIMO.Visio"
description: "Create Visio diagrams with shapes, connectors, and pages. No Visio installation required."
layout: product
product_color: "#ea580c"
install: "dotnet add package OfficeIMO.Visio"
nuget: "OfficeIMO.Visio"
docs_url: "/docs/"
api_url: "/api/"
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

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.Visio runs on Windows, Linux, and macOS. Generated files open in Microsoft Visio and other VSDX-compatible viewers.
