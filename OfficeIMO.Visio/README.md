# OfficeIMO.Visio - Visio diagrams for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Visio)](https://www.nuget.org/packages/OfficeIMO.Visio)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Visio?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Visio)

`OfficeIMO.Visio` creates, edits, inspects, validates, and exports `.vsdx`, `.vstx`, `.vssx`, `.vsdm`, `.vstm`, and `.vssm` packages without COM automation and without Microsoft Visio installed.

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
- Edits nested container topology, swimlane metadata and geometric assignment, threaded comments and authors, generated data graphics and legends, and source-preserving ShapeSheet sections and formulas.
- Provides rotation- and connector-aware resize-to-content plus deterministic topology-aware whole-page relayout for dense imported diagrams.
- Preserves opaque VBA project payloads in macro-enabled drawing, template, and stencil packages without executing or rewriting VBA.
- Exports headless PNG, JPEG, TIFF, SVG, and lossless WebP previews for proof and review workflows.
- Includes validation and quality analysis for generated and loaded diagrams, including connector-label collisions with unrelated connector paths.
- Carries caller-supplied stencil license, attribution, and unsupported-master state through shapes, catalogs, and manifests without inferring redistribution rights.

## Editing existing diagrams

`Load` materializes an editable diagram. File and stream entry points accept the
same `VisioLoadOptions`. New asynchronous calls should use the options-first
shape; token-first overloads remain available for source and binary compatibility.

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

### Loaded-diagram compatibility boundary

Loaded Open XML Visio editing covers pages, shapes, connectors, text, styles,
Shape Data, hyperlinks, layers, nested containers, swimlanes, threaded comments,
data graphics, legends, typed ShapeSheet sections/formulas, topology queries,
resize-to-content, and whole-diagram relayout. Template and stencil packages use
the same model, including page-less stencils with masters. Macro-enabled variants
retain their VBA project as an opaque bounded payload.

The boundary is deliberate: OfficeIMO does not execute VBA, evaluate arbitrary
ShapeSheet formulas, or claim native Visio layout equivalence. Typed edits retain
unmodeled ShapeSheet rows, cells, attributes, and supported opaque payloads in
their existing preservation stores. Arbitrary producer package parts outside
those stores are not advertised as editable. Ambiguous swimlane geometry is
reported instead of assigned, container cycles are rejected, and semantic
relayout keeps containers and generated adornments fixed unless the caller opts in.

### Signed diagrams

`InspectSignatures()` detects Open XML signature-origin relationships and XML signature parts. Saving a
loaded signed diagram is blocked by default because rebuilding the package would invalidate that evidence. Set
`SignatureMutationPolicy = VisioSignatureMutationPolicy.RemoveInvalidatedSignatures` only when removing the stale
signature carrier is the intended result. `SignPackageSignature(...)` and `ValidatePackageSignatures(...)` create
and cryptographically validate OPC signatures through an explicitly supplied `IOfficeSecurityProvider`.

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

When a catalog comes from an external package, make its provenance explicit:

```csharp
var options = new VisioStencilPackageLoadOptions {
    SourceLicense = "License identifier or notice supplied by the caller",
    SourceAttribution = "Required source attribution",
    IncludeUnsupportedMasters = true
};
```

Unsupported masters remain marked as unsupported even when included for inventory or migration planning. Including them does not turn them into a fully supported authoring contract or grant redistribution rights.

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

## Related packages and limits

- `OfficeIMO.Visio` generates and edits drawing, template, stencil, and macro-enabled Open XML Visio packages without requiring desktop Visio at runtime.
- External stencil packages retain their package and licensing requirements; OfficeIMO records caller-supplied provenance but never infers licensing terms.
- Use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for PowerShell workflows.
- Open Visio product work is listed in the repository [roadmap](../Docs/ROADMAP.md).

## Deeper docs

- [Repository roadmap](../Docs/ROADMAP.md)
- [Reader package family](../Docs/officeimo.reader.md)
- [Examples](../OfficeIMO.Examples)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** `System.IO.Packaging`; Microsoft BCL compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Core`. The VSDX model, builders, editing, topology, validation, and PNG/JPEG/TIFF/SVG/WebP renderers are first-party.
- **Security:** Open XML signature carriers are inspected and signed-diagram mutations fail safely without a cryptographic dependency. Signature creation and validation accept an explicit `IOfficeSecurityProvider`; `OfficeIMO.Security` is not pulled transitively.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
