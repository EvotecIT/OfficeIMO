---
title: "Visio diagram builders"
description: "Generate flowcharts, architecture diagrams, network topologies, dependency graphs, swimlanes, org charts, sequences, timelines, and generic graphs in VSDX."
meta.seo_title: "Generate Visio diagrams and VSDX files in .NET"
order: 53
---

Diagram builders turn structured application data into consistently laid-out VSDX pages. They are available for flowcharts, block diagrams, architecture, networks, network topology, dependencies, swimlanes, org charts, sequences, timelines, and generic graphs.

## Architecture diagram

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("architecture.vsdx")
    .ArchitectureDiagram("System Overview", diagram => diagram
        .Title()
        .Legend()
        .Theme(VisioStyleTheme.Technical())
        .Region("vnet", "Virtual Network", 1, 0, 3, 2)
        .Actor("users", "Users", 0, 1)
        .Gateway("gateway", "Gateway", 1, 1)
        .Service("api", "API", 2, 1)
        .Database("database", "Database", 3, 1)
        .DataFlow("users", "gateway", "HTTPS")
        .ControlFlow("gateway", "api", "route")
        .DataFlow("api", "database", "SQL"))
    .Save();
```

## Network topology

```csharp
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

## Sequence diagram

```csharp
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

## Validate the result

```csharp
IReadOnlyList<string> issues =
    VisioValidator.Validate("architecture.vsdx");

if (issues.Count > 0) {
    throw new InvalidDataException(string.Join(Environment.NewLine, issues));
}
```

Browse the [runnable Visio examples](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Examples/Visio) for data-driven identity, inventory, Kubernetes, network-segmentation, incident-runbook, and external-stencil scenarios.
