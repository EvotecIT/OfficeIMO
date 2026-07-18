# OfficeIMO.Visio - Visio diagrams for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Visio)](https://www.nuget.org/packages/OfficeIMO.Visio)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Visio?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Visio)

`OfficeIMO.Visio` creates, edits, inspects, validates, and exports `.vsdx` diagrams without COM automation and without Microsoft Visio installed.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Visio
```

## Quick start

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;

var document = VisioDocument.Create("diagram.vsdx");
document.AsFluent()
    .Info(info => info.Title("Demo").Author("OfficeIMO"))
    .Page("Page-1", page => page
        .Title("Demo Flow")
        .Rect("start", 1, 1, 2, 1, "Start")
        .Diamond("decision", 4, 1.5, 2, 2, "Decision")
        .Ellipse("end", 7, 1.5, 2, 1, "End")
        .Connect("start", "decision", VisioSide.Right, VisioSide.Left,
            connector => connector.RightAngle().ArrowEnd(EndArrow.Triangle))
        .Connect("decision", "end", VisioSide.Right, VisioSide.Left,
            connector => connector.RightAngle().ArrowEnd(EndArrow.Triangle).Label("Yes")))
    .End();
document.Save();
```

## What it does

- Creates and edits Visio pages, shapes, connectors, text, styles, Shape Data, layers, hyperlinks, containers, comments, and metadata.
- Provides fluent diagram builders for common flowchart, block, dependency, architecture, network, topology, swimlane, org chart, sequence, timeline, and generic graph scenarios.
- Supports loaded-diagram editing, shape selection, topology queries, stencil replacement/migration planning, and container maintenance.
- Exports headless PNG, JPEG, TIFF, SVG, and lossless WebP previews for proof and review workflows.
- Includes validation and quality analysis for generated and loaded diagrams.

## Editing existing diagrams

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Color = OfficeIMO.Drawing.OfficeColor;

VisioDocument.Load("operations.vsdx")
    .AsFluent()
    .ExistingPage("Operations", page => page
        .ShapesWithData("Owner", "Ops", selection => selection
            .Fill(Color.LightBlue)
            .ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean))
        .ShapesContainingText("Legacy", selection => selection
            .Text(shape => shape.Text!.Replace("Legacy", "Production", StringComparison.Ordinal))))
    .End()
    .Save("operations.updated.vsdx");
```

## Examples

The quick start shows the fluent page API. These examples show the higher-level builders and editing surfaces that belong in `OfficeIMO.Visio`.

### Flowchart builder

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("flowchart.vsdx")
    .Flowchart("Property buying flowchart", flow => flow
        .Title()
        .Layout(VisioFlowchartLayout.TwoColumnContinuation)
        .RouteBranches(laneSpacing: 0.5)
        .Start("start", "Start with an agent\nyou trust")
        .Step("consult", "Consult with agent to\ndetermine needs")
        .Decision("agreement", "Agreement?")
        .Step("contract", "Accept the contract")
        .End("close", "Close on the property")
        .Branch("agreement", "No", "consult")
        .Branch("agreement", "Yes", "contract")
        .Callout("agreement", "retry-note", "Loop back if rejected", VisioSide.Right))
    .Save();
```

### Network topology builder

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("network-topology.vsdx")
    .NetworkTopologyDiagram("Branch topology", topology => topology
        .Title()
        .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
        .Firewall("firewall", "Firewall")
        .Switch("core", "Core Switch")
        .Server("app", "App Server")
        .Database("db", "Database")
        .Workstation("finance", "Finance PC")
        .Subnet("edge", "Edge", "internet", "firewall", "core")
        .Subnet("servers", "Server Zone", "app", "db")
        .Ethernet("internet", "firewall", "WAN")
        .Trunk("firewall", "core", "uplink")
        .Trunk("core", "app", "10Gb")
        .Ethernet("app", "db"))
    .Save();
```

### Sequence diagram builder

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("sequence.vsdx")
    .SequenceDiagram("Checkout sequence", sequence => sequence
        .Title()
        .Theme(VisioStyleTheme.Fluent())
        .Actor("customer", "Customer")
        .Participant("web", "Web App")
        .Control("api", "Orders API")
        .Database("db", "Orders DB")
        .Call("customer", "web", "Checkout")
        .Call("web", "api", "POST /orders")
        .Async("api", "db", "Persist order")
        .Return("api", "web", "201 Created")
        .SelfMessage("web", "Render receipt"))
    .Save();
```

### Timeline roadmap

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("roadmap.vsdx")
    .TimelineDiagram("Product roadmap", timeline => timeline
        .Title()
        .Theme(VisioStyleTheme.Modern())
        .Range(new DateTime(2026, 1, 1), new DateTime(2026, 6, 30))
        .Span("discovery", new DateTime(2026, 1, 8), new DateTime(2026, 2, 20), "Discovery")
        .Span("build", new DateTime(2026, 2, 21), new DateTime(2026, 5, 15), "Build", lane: 1)
        .Release("preview", new DateTime(2026, 5, 20), "Public preview", VisioTimelinePlacement.Below)
        .Milestone("ga", new DateTime(2026, 6, 25), "GA"))
    .Save();
```

### Layers and Shape Data

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

var document = VisioDocument.Create("architecture.vsdx");
var page = document.AddPage("Architecture");
page.AddLayer("Infrastructure");
page.AddLayer("Annotations").Print = false;

var server = page.AddStencilShape(VisioStencils.Network.Get("server"),
    "server", 2, 5, "Server");
server.SetShapeData("Owner", "Platform", "Owner",
    VisioShapeDataType.String, "Owning support team");

page.AddToLayer("Infrastructure", server);
page.SelectWithShapeData("Owner", "Platform")
    .Fill(Color.LightBlue)
    .ShapeData("Reviewed", "Yes", "Reviewed",
        VisioShapeDataType.Boolean, "Architecture review complete");

document.Save();
```

### Headless image export

```csharp
using OfficeIMO.Visio;

var document = VisioDocument.Create("pipeline.vsdx");
var page = document.AddPage("Pipeline").Size(8, 4);
var build = page.AddProcess(1.5, 2, 1.4, 0.7, "Build");
var ship = page.AddProcess(5.5, 2, 1.4, 0.7, "Ship");
page.AddConnector(build, ship, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left)
    .EndArrow = EndArrow.Arrow;

document.SaveAsSvg("pipeline.svg", new VisioSvgSaveOptions {
    PixelsPerInch = 96,
    BackgroundColor = null
});

document.SaveAsPng("pipeline.png", new VisioPngSaveOptions {
    PixelsPerInch = 144,
    Supersampling = 3
});

OfficeImageExportResult webp = document
    .ToImage()
    .AtDpi(144)
    .AsWebp()
    .Save("pipeline.webp");

IReadOnlyList<OfficeImageExportResult> pages = document
    .ToImages()
    .AllPages()
    .AsJpeg()
    .Save("pipeline-pages");
```

## Boundaries

- `OfficeIMO.Visio` should generate and edit real `.vsdx` packages; optional desktop Visio validation belongs in examples, proof tooling, or tests.
- External stencil/package support should keep licensing and package structure explicit.
- Long assessment and roadmap notes belong in `Docs/officeimo.visio.assessment.md` and `Docs/officeimo.visio.roadmap.md`.
- PowerShell wrappers belong in [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice).

## Deeper docs

- [Visio assessment](../Docs/officeimo.visio.assessment.md)
- [Visio roadmap](../Docs/officeimo.visio.roadmap.md)
- [Document intelligence roadmap](../Docs/officeimo.document-intelligence-roadmap.md)
- [Examples](../OfficeIMO.Examples)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** `System.IO.Packaging`; Microsoft BCL compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Drawing`. The VSDX model, builders, editing, topology, validation, and PNG/JPEG/TIFF/SVG/WebP renderers are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
