---
title: "Visio stencils and shape catalogs"
description: "Use OfficeIMO shape catalogs, installed Microsoft Visio stencils, or external VSSX and VSTX packs while keeping master and licensing choices explicit."
meta.seo_title: "Use VSSX and VSTX Visio stencils from .NET"
order: 55
---

OfficeIMO.Visio can place shapes from generated catalogs, installed Visio stencil directories, and explicitly supplied `.vssx` or `.vstx` packages. The package retains the relationship between a shape and its master so a generated diagram can be edited and migrated later.

## Use a built-in catalog

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;

var document = VisioDocument.Create("network.vsdx");
var page = document.AddPage("Network");

var server = page.AddStencilShape(
    VisioStencils.Network.Get("server"),
    "server",
    2,
    4,
    "Application server");

var database = page.AddStencilShape(
    VisioStencils.Network.Get("database"),
    "database",
    6,
    4,
    "Orders database");

page.AddConnector(
    server,
    database,
    ConnectorKind.RightAngle,
    VisioSide.Right,
    VisioSide.Left);

document.Save();
```

## External and installed stencils

Use an explicit stencil path when the application owns or is licensed to use a VSSX/VSTX pack. Discovery helpers can inspect installed Visio stencil locations on Windows, while server and cross-platform applications can point at a deployed stencil repository.

The [external stencil examples](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Examples/Visio) show catalog loading, master inspection, placement, and validation. They do not bundle third-party stencil packs.

## Keep mappings stable

1. Identify masters by a stable catalog key rather than display order.
2. Record the source stencil and expected master identity.
3. Keep application data separate from shape-placement choices.
4. Validate the generated VSDX after a stencil-pack upgrade.
5. Use migration planning before replacing masters in existing diagrams.

Continue with [editing and validation](/docs/visio/editing-and-validation/) for updating loaded shapes and checking round trips.
