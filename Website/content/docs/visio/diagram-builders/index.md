---
title: "Generate Visio Diagrams with Builders"
description: "Generate architecture, flowchart, network topology, and sequence diagrams as VSDX files from structured data with OfficeIMO.Visio builders."
meta.seo_title: "Generate Visio diagrams and VSDX files in .NET"
order: 53
packages:
  - OfficeIMO.Visio
search_terms:
  - generate visio diagram
  - flowchart vsdx
  - network diagram
  - architecture diagram
  - sequence diagram
verified: 2026-11-14
---

Diagram builders in `OfficeIMO.Visio.Diagrams` turn application data into laid-out VSDX pages. Each builder is an extension method on `VisioDocument` that adds one page: `ArchitectureDiagram`, `Flowchart`, `NetworkDiagram`, `NetworkTopologyDiagram`, `DependencyDiagram`, `BlockDiagram`, `SwimlaneDiagram`, `OrgChartDiagram`, `SequenceDiagram`, `TimelineDiagram`, and `GraphDiagram`. You give nodes and links by ID; the builder places shapes and routes connectors.

## Generate an architecture diagram

Region, actor, and service positions are grid cells (column, row).

### C#

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

using VisioDocument document = VisioDocument.Create("architecture.vsdx");

document.ArchitectureDiagram("System Overview", diagram => diagram
    .Title()
    .Legend()
    .Theme(VisioStyleTheme.Technical())
    .PageSize(14, 8.5)
    .Region("vnet", "Virtual Network", 1, 0, 3, 2)
    .Actor("users", "Users", 0, 1)
    .Gateway("gateway", "Gateway", 1, 1)
    .Service("api", "API", 2, 1)
    .Database("database", "Database", 3, 1)
    .DataFlow("users", "gateway", "HTTPS")
    .ControlFlow("gateway", "api", "route")
    .DataFlow("api", "database", "SQL"));

document.Save();
```

## Generate a flowchart

### C#

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

using VisioDocument document = VisioDocument.Create("approval.vsdx");

document.Flowchart("Approval", flow => flow
    .Start("start", "Request received")
    .Step("review", "Review request")
    .Decision("approved", "Approved?")
    .Step("fulfil", "Fulfill request")
    .End("done", "Done")
    .Branch("approved", "No", "review"));

document.Pages[0].FitToContent(0.7, 0.55);
document.Save();
```

## Generate a network topology

### C#

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

using VisioDocument document = VisioDocument.Create("network-topology.vsdx");

document.NetworkTopologyDiagram("Branch topology", topology => topology
    .Title()
    .Theme(VisioStyleTheme.Technical())
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
    .Ethernet("app", "db")
    .Ethernet("core", "finance"));

document.Save();
```

## Generate a sequence diagram

### C#

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

using VisioDocument document = VisioDocument.Create("sequence.vsdx");

document.SequenceDiagram("Checkout", sequence => sequence
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
    .SelfMessage("web", "Render receipt"));

document.Save();
```

## Check the generated file

`Validate` checks package and diagram rules that OfficeIMO knows about. It returns an empty list when nothing is wrong.

### C#

```csharp
using OfficeIMO.Visio;

IReadOnlyList<string> issues = VisioValidator.Validate("architecture.vsdx");

foreach (string issue in issues) {
    Console.WriteLine(issue);
}
```

More data-driven examples (identity, inventory, Kubernetes, network segmentation, incident runbooks) are in [OfficeIMO.Examples/Visio](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Examples/Visio).
