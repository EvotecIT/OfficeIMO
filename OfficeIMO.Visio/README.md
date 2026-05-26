# OfficeIMO.Visio — .NET Visio Utilities

OfficeIMO.Visio provides helpers for creating and editing .vsdx drawings with Open XML.

- Targets: netstandard2.0, net472 (Windows), net8.0, net10.0
- License: MIT
- NuGet: `OfficeIMO.Visio`
- Dependencies: OfficeIMO.Drawing, System.IO.Packaging (Windows), Microsoft.Bcl.AsyncInterfaces (net472)

## Install

```powershell
dotnet add package OfficeIMO.Visio
```

## Quick sample (fluent)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;

var vsd = VisioDocument.Create("diagram.vsdx");
vsd.AsFluent()
   .Info(i => i.Title("Demo").Author("You"))
   .Page("Page-1", p => p
       .Title("Demo Flow")
       .Rect("start", 1, 1, 2, 1, "Start")
       .Diamond("decision", 4, 1.5, 2, 2, "Decision")
       .Ellipse("end", 7, 1.5, 2, 1, "End")
       .Connect("start", "decision", VisioSide.Right, VisioSide.Left,
           c => c.RightAngle().ArrowEnd(EndArrow.Triangle))
       .Connect("decision", "end", VisioSide.Right, VisioSide.Left,
           c => c.RightAngle().ArrowEnd(EndArrow.Triangle).Label("Yes")))
   .End();
vsd.Save();
```

## Quick sample (diagram builder)

```csharp
using System;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("flowchart.vsdx")
    .Flowchart("Property buying Flowchart", flow => flow
        .Title()
        .Layout(VisioFlowchartLayout.TwoColumnContinuation)
        .RouteBranches(laneSpacing: 0.5)
        .Start("start", "Start with an agent\nyou trust")
        .Step("consult", "Consult with agent to\ndetermine your property\nwants and needs")
        .Step("market", "With agent, analyze\nmarket to choose\nproperties of interest")
        .OffPage("jump", "A")
        .Continue("resume", "A")
        .Step("offer", "Select ideal property\nand write offer to\npurchase")
        .Decision("agreement", "Negotiate\n& Counteroffer:\nAgreement?")
        .Step("contract", "Accept the contract")
        .End("close", "Close on the\nproperty")
        .Branch("agreement", "No", "market"))
    .Save();
```

The diagram builder creates normal Visio pages, semantic flowchart shapes,
masters, side-glued connectors, labels, deterministic layouts, and routed
branch/loop connectors. It is the first high-level authoring layer above the
lower-level page/shape APIs.

## Quick sample (block diagram builder)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("block-diagram.vsdx")
    .BlockDiagram("Block Diagram", diagram => diagram
        .Title()
        .Legend()
        .Region("processor", "Processor", 1, 2, 2, 2)
        .Block("input", "Input Device", 0, 2)
        .EmphasisBlock("memory", "Memory Unit", 1, 2)
        .Block("storage", "Secondary\nStorage", 1, 0, VisioBlockShapeKind.Data)
        .Block("control", "Control Unit", 1, 3)
        .Block("alu", "Arithmetic &\nLogic Unit", 1, 4)
        .Block("output", "Output Device", 3, 2)
        .DataFlow("input", "memory")
        .DataFlow("memory", "output")
        .ControlFlow("control", "output", "Control Flow"))
    .Save();
```

The block diagram builder creates grid-positioned blocks, light background
regions, solid data-flow connectors, dashed control-flow connectors, labels,
optional presentation titles/legends, and master-backed Visio shapes.

## Quick sample (dependency diagram builder)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("dependencies.vsdx")
    .DependencyDiagram("Service Dependencies", diagram => diagram
        .Theme(VisioStyleTheme.Fluent())
        .External("users", "Users")
        .Component("web", "Web App")
        .Component("api", "API")
        .Decision("policy", "Policy")
        .Data("database", "Database")
        .DependsOn("users", "web", "HTTPS")
        .DependsOn("web", "api")
        .ControlDependency("api", "policy", "Authorize")
        .DataDependency("api", "database", "SQL"))
    .EnsureVisualQuality(new VisioDiagramQualityOptions {
        CheckConnectorShapeIntersections = false,
        CheckConnectorLabelShapeOverlaps = false
    })
    .Save();
```

The dependency diagram builder creates deterministic layered DAG layouts from
nodes and directed relationships. It automatically grows the page, places
component/data/external/decision nodes, routes dependencies, and rejects cycles.

## Quick sample (architecture diagram builder)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("architecture.vsdx")
    .ArchitectureDiagram("Jenkins on Azure", diagram => diagram
        .Title()
        .Legend()
        .Theme(VisioStyleTheme.Technical())
        .Region("vnet", "Virtual Network", 1, 0, 4, 3)
        .Region("subnet", "Build Subnet", 1, 1, 4, 2)
        .Actor("users", "Users", 0, 1)
        .Gateway("public-ip", "Public IP", 1, 1)
        .Service("jenkins", "Jenkins Server", 2, 1)
        .Compute("agent", "Build Agent", 3, 1)
        .Database("data", "Data", 2, 2)
        .Storage("artifacts", "Artifacts", 4, 2)
        .Security("vault", "Key Vault", 2, 0)
        .DataFlow("users", "public-ip", "HTTPS")
        .DataFlow("public-ip", "jenkins", "route")
        .ControlFlow("jenkins", "agent", "scale")
        .Dependency("jenkins", "vault", "secrets")
        .Callout("jenkins", "scale-note", "Scale agents on demand", 7.8, 6.3))
    .Save();
```

The architecture builder creates dependency-free cloud/infrastructure diagrams
with semantic components, background regions, routed data/control/dependency
connectors, labels, callouts, and the reusable `Technical` theme.

## Quick sample (network diagram builder)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("network.vsdx")
    .NetworkDiagram("Branch Network", network => network
        .Theme(VisioStyleTheme.Technical())
        .Zone("perimeter", "Perimeter", 0, 0, 3, 1)
        .Zone("servers", "Server Zone", 3, 0, 3, 1)
        .Zone("clients", "Client LAN", 1, 2, 5, 1)
        .Internet("internet", "Internet", 0, 0)
        .Firewall("firewall", "Firewall", 1, 0)
        .Switch("core", "Core Switch", 2, 0)
        .Server("app", "App Server", 3, 0)
        .Database("db", "Database", 4, 0)
        .Workstation("pc1", "Finance PC", 1, 2)
        .Workstation("pc2", "Support PC", 2, 2)
        .Printer("printer", "Printer", 3, 2)
        .Ethernet("internet", "firewall", "WAN")
        .Trunk("firewall", "core", "uplink")
        .Trunk("core", "app", "10Gb")
        .Ethernet("app", "db")
        .Ethernet("core", "pc2")
        .Ethernet("pc2", "printer"))
    .Save();
```

The network builder creates dependency-free network maps with zones, typed
devices, routed Ethernet/trunk/wireless/management links, and optional legends.

## Quick sample (network topology diagram builder)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("network-topology.vsdx")
    .NetworkTopologyDiagram("Branch Topology", topology => topology
        .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
        .Firewall("firewall", "Firewall")
        .Switch("core", "Core Switch")
        .Server("app", "App Server")
        .Database("db", "Database")
        .Workstation("finance", "Finance PC")
        .Workstation("support", "Support PC")
        .Printer("printer", "Printer")
        .Subnet("edge", "Edge", "internet", "firewall", "core")
        .Subnet("servers", "Server Zone", "app", "db")
        .Subnet("clients", "Client LAN", "finance", "support", "printer")
        .Ethernet("internet", "firewall", "WAN")
        .Trunk("firewall", "core", "uplink")
        .Trunk("core", "app", "10Gb")
        .Ethernet("app", "db")
        .Ethernet("core", "finance")
        .Ethernet("core", "support")
        .Ethernet("support", "printer"))
    .Save();
```

The topology builder is the graph-first network API: users describe devices
and links, then OfficeIMO derives deterministic layers, grows the page when
needed, adds subnet/background zones around selected devices, routes links,
and keeps mesh/cycle links valid.

## Quick sample (swimlane diagram builder)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("swimlane.vsdx")
    .SwimlaneDiagram("Order Fulfillment", swim => swim
        .Theme(VisioStyleTheme.Modern())
        .Lane("customer", "Customer")
        .Lane("sales", "Sales")
        .Lane("ops", "Operations")
        .Phase("request", "Request")
        .Phase("review", "Review")
        .Phase("approval", "Approval")
        .Phase("fulfill", "Fulfill")
        .Start("start", "Submit order", "customer", "request")
        .Step("qualify", "Qualify order", "sales", "review")
        .Decision("approved", "Approved?", "sales", "approval")
        .Step("revise", "Revise request", "customer", "approval")
        .Step("pick", "Pick items", "ops", "approval")
        .Data("invoice", "Create invoice", "sales", "fulfill")
        .End("ship", "Ship order", "ops", "fulfill")
        .Flow("start", "qualify", "handoff")
        .Flow("qualify", "approved")
        .Exception("approved", "revise", "no")
        .Handoff("approved", "pick", "yes")
        .Flow("pick", "invoice")
        .Flow("invoice", "ship"))
    .Save();
```

The swimlane builder creates editable role lanes, phase headers, semantic
activities, labeled flows, dashed exception paths, deterministic routing, and
automatic stacking when more than one activity lands in the same lane/phase
cell. It does not require Visio templates at runtime.

## Quick sample (org chart builder)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("org-chart.vsdx")
    .OrgChartDiagram("Leadership", org => org
        .Theme(VisioStyleTheme.Modern())
        .Root("ceo", "Marta Nowak", "Chief Executive Officer")
        .Assistant("ea", "Eli Green", "Executive Assistant", "ceo")
        .Manager("cto", "Alex Chen", "Chief Technology Officer", "ceo")
        .Manager("coo", "Sam Rivera", "Chief Operating Officer", "ceo")
        .Manager("cfo", "Priya Shah", "Chief Financial Officer", "ceo")
        .TeamBand("engineering", "Engineering", "cto")
        .TeamBand("operations", "Operations", "coo")
        .Position("platform", "Nina Patel", "Platform Lead", "cto", "engineering")
        .Position("security", "Owen Brooks", "Security Lead", "cto", "engineering")
        .Vacancy("sre", "Open SRE Role", "coo", "operations")
        .External("advisor", "Taylor Reed", "Advisor", "cfo"))
    .Save();
```

The org chart builder creates editable hierarchy cards, assistant placements,
team bands, vacancies, external roles, and routed reporting lines from semantic
relationships.

## Reusable style themes

`VisioStyleTheme` gives diagrams and later editing passes a shared set of
shape, connector, and readable text styles. The built-in presets are `Modern`,
`Office`, `Fluent`, `Technical`, `Minimal`, `Dark`, and `Print`.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

var theme = VisioStyleTheme.Minimal();
var dark = VisioStyleTheme.Dark();

var doc = VisioDocument.Create("styled.vsdx")
    .Flowchart("Styled Approval Flow", flow => flow
        .Theme(theme)
        .Start("start", "Request received")
        .Step("review", "Review request")
        .Decision("approved", "Approved?")
        .End("done", "Done"));

var page = doc.Pages[0];
page.SelectByMaster("Decision").Style(theme.Decision);
page.SelectConnectedConnectors(page.FindShapeById("approved")!)
    .Style(theme.ControlConnector);
page.FitToContent(0.6, 0.45);
doc.Save();

VisioDocument.Create("dark-styled.vsdx")
    .BlockDiagram("Dark System Blocks", diagram => diagram
        .Theme(dark)
        .Region("zone", "Processing Zone", 0, 0, 3, 1)
        .Block("input", "Input", 0, 0)
        .EmphasisBlock("processor", "Processor", 1, 0)
        .Block("output", "Output", 2, 0)
        .DataFlow("input", "processor")
        .ControlFlow("processor", "output", "control"))
    .Save();
```

## Connector routing

Connectors can stay dynamic for Visio-managed rerouting, or they can be pinned
to deterministic OfficeIMO-generated orthogonal routes and explicit waypoints.
This is useful for readable flowcharts, architecture diagrams, and edited
documents where a few important lines must avoid crossing the main content.
OfficeIMO-authored explicit waypoint routes also load back into the connector
model, so routed diagrams can be edited and saved again without losing the
route semantics.
Pages can also set native Visio routing defaults for connectors that do not
carry local routing or line-jump settings, plus placement and layout-grid policy
used by Visio's Re-Layout Page commands.

```csharp
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;

var doc = VisioDocument.Create("routes.vsdx");
var page = doc.AddPage("Routes");
page.PlacementStyle = VisioPlacementStyle.HierarchyLeftToRightMiddle;
page.PlacementDepth = VisioPlacementDepth.Medium;
page.PlacementFlip = VisioPlacementFlip.Horizontal | VisioPlacementFlip.Rotate90;
page.MoveShapesAwayOnDrop = true;
page.ResizePageToFitLayout = true;
page.EnableLayoutGrid = true;
page.SetLayoutGridSizing(1.2, 0.45);
page.ConnectorRouteStyle = VisioPageRouteStyle.FlowchartTopToBottom;
page.ConnectorRouteAppearance = VisioLineRouteExtension.Straight;
page.LineJumpStyle = VisioLineJumpStyle.Gap;
page.LineJumpCode = VisioLineJumpCode.DisplayOrder;
page.HorizontalLineJumpDirection = VisioHorizontalLineJumpDirection.Up;
page.VerticalLineJumpDirection = VisioVerticalLineJumpDirection.Right;
page.SetConnectorSpacing(0.25, 0.45);

var source = page.AddStencilShape(VisioStencils.Flowchart.Get("process"),
    "source", 2, 5, "Source");
var target = page.AddStencilShape(VisioStencils.Flowchart.Get("process"),
    "target", 7, 3, "Target");
target.PlacementStyle = VisioPlacementStyle.HierarchyLeftToRightMiddle;
target.PlacementFlip = VisioPlacementFlip.Horizontal | VisioPlacementFlip.Rotate90;
target.PlowCode = VisioShapePlowCode.Always;
target.AllowHorizontalConnectorRoutingThrough = false;
target.AllowVerticalConnectorRoutingThrough = false;

VisioConnector route = page.AddConnector(source, target, ConnectorKind.Dynamic,
        VisioSide.Right, VisioSide.Left);
route.RouteStyle = VisioPageRouteStyle.FlowchartLeftToRight;
route.RouteAppearance = VisioLineRouteExtension.Curved;
route.LineJumpStyle = VisioLineJumpStyle.Square;
route.LineJumpCode = VisioConnectorLineJumpCode.Always;
route.HorizontalJumpDirection = VisioHorizontalLineJumpDirection.Up;
route.VerticalJumpDirection = VisioVerticalLineJumpDirection.Right;
route.RerouteBehavior = VisioConnectorRerouteBehavior.OnCrossover;
route
    .RouteOrthogonal(VisioConnectorRouteStyle.HorizontalThenVertical)
    .PlaceLabel(0.65, offsetY: 0.18)
    .ApplyStyle(VisioStyleTheme.Modern().Connector);

page.SelectConnectedConnectors(source)
    .RouteThrough(VisioConnectorWaypoint.At(4.5, 5),
        VisioConnectorWaypoint.At(4.5, 3))
    .Label("handoff")
    .LabelPosition(0.6, offsetX: 0.15);

doc.Save();
```

## Timeline roadmaps

The timeline builder creates date-scaled roadmap diagrams with milestone
semantics, above/below placement, stacked labels, and span lanes. It is useful
for release plans, migration schedules, project phases, and executive roadmap
views where the author should provide dates, not hand-place every marker.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("roadmap.vsdx")
    .TimelineDiagram("Product Roadmap", timeline => timeline
        .Theme(VisioStyleTheme.Modern())
        .Range(new DateTime(2026, 1, 1), new DateTime(2026, 6, 30))
        .Span("discovery", new DateTime(2026, 1, 8), new DateTime(2026, 2, 20), "Discovery")
        .Span("build", new DateTime(2026, 2, 21), new DateTime(2026, 5, 15), "Build", lane: 1)
        .Release("preview", new DateTime(2026, 5, 20), "Public preview", VisioTimelinePlacement.Below)
        .Milestone("ga", new DateTime(2026, 6, 25), "GA"))
    .Save();
```

## Visual quality checks and gallery output

Package validation proves the `.vsdx` structure is sound. Visual quality checks
catch common diagram problems before a human opens Visio: shapes outside the
page, overlapping shapes, routed connectors crossing unrelated shapes, and
connector labels placed off-page, over unrelated shapes, or on top of each
other.

```csharp
using OfficeIMO.Visio;

var results = VisioGallery.Create("gallery");
foreach (var result in results) {
    if (!result.IsClean) {
        foreach (var issue in result.QualityIssues) {
            Console.WriteLine(issue);
        }
    }
}

var proofOptions = new VisioGalleryOptions {
    ValidateWithVisioDesktop = true,
    DesktopValidationOptions = VisioDesktopValidationOptions.RoundTripWithSvg()
};
var proofResults = VisioGallery.Create("gallery-proof", proofOptions);

var issues = doc.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
    RequireConnectorLabels = false,
    CheckConnectorLabelOverlaps = true,
    CheckConnectorLabelShapeOverlaps = true
});
var report = doc.GetVisualQualityReport();
doc.EnsureVisualQuality(minimumSeverity: VisioDiagramQualityIssueSeverity.Warning);

page.ResolveConnectorLabelOverlaps();
doc.PolishDiagrams();
```

`EnsureVisualQuality(...)` throws `VisioDiagramQualityException` with the
blocking issues, which makes it practical to use generated diagrams in tests or
CI without writing custom issue-loop code.

## Native stencil catalogs

Built-in stencil catalogs give you reusable, searchable shape definitions while
still generating masters from OfficeIMO code. They are not `.vssx` or `.vsdx`
runtime dependencies. Use `Get(...)` for exact known shapes and `Search(...)`
or `InCategory(...)` when you want user-friendly discovery by id, name, master,
category, keyword, alias, or tag. Each stencil also carries `IconNameU` preview
metadata for palette and picker UIs.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;

var doc = VisioDocument.Create("stencils.vsdx");
var page = doc.AddPage("Catalog");

var process = page.AddStencilShape(VisioStencils.Flowchart.Get("process"),
    "receive", 2, 4, "Receive request");
var decision = page.AddStencilShape(VisioStencils.Flowchart, "branch",
    "approved", 5, 4, "Approved?");
var switchShape = page.AddStencilShape("net.switch", "switch", 8, 6, "Switch");
var dataStore = VisioStencils.All.Search("data-store").First();
var networkShapes = VisioStencils.All.InCategory("Network");
var custom = VisioStencilCatalog.Create("Custom Infrastructure", catalog => catalog
    .Add("custom.cache", "Cache", "Process", "Infrastructure", 1.8, 0.9, "redis")
    .AddWithMetadata("custom.archive", "Object Archive", "Data", "Infrastructure",
        1.8, 0.9,
        keywords: new[] { "blob" },
        aliases: new[] { "object-store" },
        tags: new[] { "cloud", "storage" },
        iconNameU: "Data"));
var cache = page.AddStencilShape(custom, "redis", "cache", 8, 4, "Cache");
var packageCatalog = VisioStencilPackageCatalog.Load("network.vssx",
    new VisioStencilPackageLoadOptions {
        Category = "Network",
        MasterNames = new[] { "Server", "rId4", "database-cylinder" },
        IncludeUnsupportedMasters = false
    });
custom.Save("infrastructure.officeimo-visio-stencils.xml");
var reusable = VisioStencilCatalog.Load("infrastructure.officeimo-visio-stencils.xml");

page.AddConnector(process, decision, ConnectorKind.Dynamic,
    VisioSide.Right, VisioSide.Left);
doc.Save();
```

`VisioStencilPackageCatalog.Load(...)` reads master metadata from `.vsdx`,
`.vssx`, and `.vstx` packages. It does not use those files as runtime templates;
by default it only exposes masters that OfficeIMO can generate natively. The
`MasterNames` filter can target the universal name, visible name, relationship id,
numeric id, or normalized slug discovered in the package. Set
`IncludeUnsupportedMasters` only when a generic generated placeholder is useful
for discovery, migration tooling, or keeping a learned palette placeable without
shipping the source stencil or template.

`VisioStencilCatalog.Save(...)` and `VisioStencilCatalog.Load(...)` persist
OfficeIMO-native catalog metadata as a small XML manifest. This is useful for
reusable first-party or application-specific palettes without requiring Visio
stencil packages at runtime.

## Query and selection editing

Query helpers let you edit diagrams by meaning instead of by page indexes. They
work with generated and loaded shapes, including nested group children.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

var doc = VisioDocument.Create("editing.vsdx");
var page = doc.AddPage("Ownership");

var intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"),
    "intake", 2, 5, "Receive");
var review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"),
    "review", 5, 5, "Review");
var decision = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"),
    "approved", 8, 5, "Approved?");

page.AddConnector(intake, review, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
page.AddConnector(review, decision, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);

intake.Data["Owner"] = "Ops";
review.Data["Owner"] = "Ops";

page.SelectWithData("Owner", "Ops")
    .Fill(Color.LightBlue)
    .Stroke(Color.DodgerBlue, 0.02)
    .Duplicate(1.5, -0.75);

page.SelectOutgoingConnectors(review)
    .LineColor(Color.DodgerBlue)
    .EndArrow(EndArrow.Triangle);

doc.Save();
```

Selection duplication remaps copied shape identifiers and duplicates only the
connectors whose endpoints are both inside the copied selection. Shape styling,
layers, hyperlinks, User cells, typed Shape Data, protection, layout hints, and
connector routing metadata move with the copy.

Whole pages can be duplicated as well. The copy receives fresh shape and
connector IDs while keeping page settings, layers, background-page linkage,
shape metadata, internal connectors, labels, and explicit routes:

```csharp
VisioPage reviewPage = page.Duplicate("Review copy");

VisioPage independentReviewPage = page.Duplicate(new VisioPageDuplicationOptions {
    Name = "Review copy",
    DuplicateBackgroundPage = true,
    BackgroundPageName = "Review background copy"
});
```

Existing shapes can also be retargeted to a different generated master without
losing their editing metadata or connector endpoints:

```csharp
page.SelectByMaster("Process")
    .ReplaceMaster(VisioStencils.Flowchart.Get("decision"), resizeToMaster: true);

page.ReplaceMaster(archiveShape, "Data");
```

## Layers

Pages support Visio-native layers. Shapes and connectors can belong to one or
more layers; OfficeIMO writes the page `Layer` section and the shape
`LayerMember` cells directly, so the document opens in Visio with editable
layer membership.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

var doc = VisioDocument.Create("layers.vsdx");
var page = doc.AddPage("Layered");
page.AddLayer("Infrastructure");
page.AddLayer("Annotations").Print = false;

var server = page.AddStencilShape(VisioStencils.Network.Get("server"),
    "server", 2, 5, "Server");
var note = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"),
    "note", 5, 5, "Internal note");

page.AddToLayer("Infrastructure", server)
    .AddToLayer("Annotations", note);

page.SelectLayer("Infrastructure")
    .Stroke(Color.DodgerBlue, 0.02);

doc.Save();
```

## Hyperlinks

Shapes and connectors can carry native Visio hyperlink rows. OfficeIMO writes
the ShapeSheet `Hyperlink` section directly and loads it back into typed
`VisioHyperlink` objects, while preserving unknown hyperlink cells from files
created elsewhere.

```csharp
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

var doc = VisioDocument.Create("hyperlinks.vsdx");
var page = doc.AddPage("Linked");

var portal = page.AddRectangle(2, 5, 2, 1, "Portal");
var api = page.AddRectangle(5, 5, 2, 1, "API");

portal.AddHyperlink("https://github.com/EvotecIT/OfficeIMO", "Repository");
var connector = page.AddConnector(portal, api, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
connector.AddHyperlink("https://example.org/openapi.json", "API contract");

page.SelectWithHyperlinks()
    .Fill(Color.LightYellow)
    .Stroke(Color.DodgerBlue, 0.02);
page.SelectConnectorsWithHyperlinks()
    .EndArrow(EndArrow.Triangle);

doc.Save();
```

## Shape Data

Shapes support typed Visio Shape Data rows in the ShapeSheet `Prop` section.
The simple `Data` dictionary still works, while `SetShapeData` lets you keep
labels, prompts, types, formats, sort keys, and other metadata visible in
Visio's Shape Data window.

```csharp
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

var doc = VisioDocument.Create("shape-data.vsdx");
var page = doc.AddPage("Shape Data");

var api = page.AddRectangle(2.5, 4, 2.2, 1, "API");
api.SetShapeData("Owner", "Platform", "Owner",
    VisioShapeDataType.String, "Owning support team");
api.SetShapeData("MonthlyCost", "1250", "Monthly cost",
    VisioShapeDataType.Currency, "Estimated monthly cost", "$#,##0");

var database = page.AddRectangle(6, 4, 2.2, 1, "Database");
database.SetShapeData("Owner", "Data", "Owner",
    VisioShapeDataType.String, "Owning support team");

page.SelectWithShapeData("Owner", "Platform")
    .Fill(Color.LightBlue)
    .ShapeData("Reviewed", "Yes", "Reviewed",
        VisioShapeDataType.Boolean, "Architecture review complete");

doc.Save();
```

## Page settings

Pages expose common print and page-management cells without requiring raw
ShapeSheet XML: margins, print orientation, page replacement/duplication locks,
drawing size behavior, automatic page resizing, shape splitting, and whether a
page is visible in Visio page lists.

```csharp
using OfficeIMO.Visio;

var doc = VisioDocument.Create("page-settings.vsdx");
var page = doc.AddPage("Print ready", 11, 8.5);
page.SetMargins(0.4, 0.5, 0.6, 0.7);
page.PrintOrientation = VisioPagePrintOrientation.Landscape;
page.PageLockReplace = true;
page.DrawingSizeType = VisioDrawingSizeType.Custom;
page.AutoResizeDrawing = false;
page.AllowShapeSplitting = false;
page.UiVisibility = VisioPageUiVisibility.Normal;
page.AddRectangle(5.5, 4.25, 2.4, 1, "Print-ready page");

doc.Save();
```

## Background pages

Reusable Visio background pages can hold title bands, watermarks, legends, page
frames, or locked diagram furniture once, then foreground pages can reference
them with native `BackPage` links.

```csharp
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

var doc = VisioDocument.Create("background-pages.vsdx");

var background = doc.AddBackgroundPage("Brand background", 11, 8.5);
background.AddRectangle(5.5, 8.05, 10.5, 0.45, "OfficeIMO generated")
    .Protect(p => p.Size().Position().Text().Selection())
    .FillColor = Color.LightBlue;

var architecture = doc.AddPage("Architecture", 11, 8.5);
architecture.SetBackgroundPage(background);
architecture.AddRectangle(3.5, 4.8, 2.2, 1, "API");
architecture.AddRectangle(7.5, 4.8, 2.2, 1, "Worker");

var operations = doc.AddPage("Operations", 11, 8.5);
operations.SetBackgroundPage(background);
operations.AddRectangle(5.5, 4.8, 2.2, 1, "Runbook");

doc.Save();
```

## Shape protection

Shapes and connectors expose native Visio `Lock*` ShapeSheet cells for diagrams
that should open cleanly in Visio but protect generated scaffolding and routed
connectors from accidental edits. Protection round-trips from existing `.vsdx`
files and works with selections.

```csharp
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

var doc = VisioDocument.Create("protected-diagram.vsdx");
var page = doc.AddPage("Protected Diagram");

var background = page.AddRectangle(4.25, 3, 7.5, 4.8, "Generated zone");
background.FillColor = Color.LightCyan;
background.Protect(p => p.Size().Position().Selection().Formatting());

var api = page.AddRectangle(3, 3.8, 2, 1, "API");
api.SetShapeData("Owner", "Platform");
var db = page.AddRectangle(6, 3.8, 2, 1, "Database");
var link = page.AddConnector(api, db, ConnectorKind.Dynamic,
    VisioSide.Right, VisioSide.Left);
link.Label = "read";
link.Protect(p => p.Endpoints().Text().Deletion());

page.SelectWithData("Owner", "Platform")
    .Protect(p => p.Text().Deletion())
    .Fill(Color.LightYellow);

page.SelectConnectorsWithProtection()
    .Protect(p => p.Formatting());

doc.Save();
```

## Containers and User cells

Pages can create Visio-native containers around existing shapes. OfficeIMO
writes the ShapeSheet `User` section and `Relationships` cells directly, so
containers open as semantic Visio structures rather than as decorative boxes.
Generic User cells are also available for custom ShapeSheet metadata.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

var doc = VisioDocument.Create("containers.vsdx");
var page = doc.AddPage("Application");

var api = page.AddStencilShape(VisioStencils.Network.Get("server"),
    "api", 3, 5.5, "API");
var worker = page.AddStencilShape(VisioStencils.Network.Get("server"),
    "worker", 6, 5.5, "Worker");

var tier = page.AddContainer("app-tier", "Application tier",
    new[] { api, worker },
    new VisioContainerOptions {
        Margin = 0.35,
        FillColor = Color.LightCyan,
        LineColor = Color.DodgerBlue
    });

tier.SetUserCell("OfficeIMO.Role", "Tier", "STR", prompt: "Semantic role");

page.SelectContainers()
    .Stroke(Color.DodgerBlue, 0.02);
page.SelectWithUserCell("OfficeIMO.Role", "Tier")
    .UserCell("OfficeIMO.Reviewed", "Yes", "STR");

doc.Save();
```

## Callouts and annotations

Pages can add semantic callouts with leader connectors. OfficeIMO writes normal
editable Visio shapes and connectors, plus User cells that make callouts easy
to find again after loading.

```csharp
using OfficeIMO.Visio;

var api = page.AddProcess(4, 4.5, 2, 1, "API");
var note = page.AddCallout(api, "api-note", "Check retry policy", 7.5, 6,
    new VisioCalloutOptions {
        Width = 2.4,
        Height = 0.8,
        RouteOffset = 0.15
    });

page.SelectCallouts()
    .LockPosition();

doc.Save();
```

## Text styling

Text blocks can be styled with a reusable object or applied to selections. The
same style model works for shape text and connector labels. Styles are saved
into Visio ShapeSheet text block cells plus `Char` and `Para` sections, so text
remains editable in Visio.

```csharp
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

var textStyle = new VisioTextStyle {
    FontFamily = "Aptos",
    Color = Color.FromRgb(0x33, 0x66, 0x99),
    Size = 12,
    Bold = true,
    HorizontalAlignment = VisioTextHorizontalAlignment.Center,
    VerticalAlignment = VisioTextVerticalAlignment.Middle,
    LeftMargin = 0.08,
    RightMargin = 0.08,
    TextPinY = -0.2,
    TextHeight = 0.4,
    BackgroundColor = Color.LightYellow,
    BackgroundTransparency = 20
};

page.AddProcess(2, 4, 2.5, 1, "Approve")
    .ApplyTextStyle(textStyle);

page.SelectWithData("Lane", "Finance")
    .ApplyTextStyle(textStyle);

var connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
connector.Label = "Approved";
connector
    .PlaceLabel(0.55, offsetY: 0.18)
    .ApplyTextStyle(new VisioTextStyle {
        FontFamily = "Aptos",
        Color = Color.DodgerBlue,
        Size = 9,
        Bold = true,
        HorizontalAlignment = VisioTextHorizontalAlignment.Center,
        BackgroundColor = Color.White,
        BackgroundTransparency = 0
    });

page.SelectConnectedConnectors(source)
    .ApplyTextStyle(textStyle);

doc.Save();
```

## Layout cleanup helpers

Selections can also be aligned, distributed, resized to text, centered, and fit
to page bounds. Page fitting and centering include explicit connector waypoints
and connector label boxes, so routed labels do not get clipped. Text sizing uses
the deterministic `OfficeIMO.Drawing` measurement engine, so it works without
system font APIs. Connector labels can also be nudged deterministically away
from page edges, unrelated shapes, and labels that were already placed.

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Visio;

page.SelectWithData("Owner", "Ops")
    .ResizeToText(new OfficeFontInfo("Calibri", 11))
    .RelayoutAsGrid(columns: 2, horizontalSpacing: 0.4, verticalSpacing: 0.3)
    .Align(VisioVerticalAlignment.Middle);

page.SelectByMaster("Decision")
    .Align(VisioHorizontalAlignment.Center);

page.SelectConnectedConnectors(page.FindShapeById("approved")!)
    .ApplyTextStyle(new VisioTextStyle { FontFamily = "Calibri", Size = 9 })
    .ResizeLabelsToText(maximumWidth: 1.6);

page.PolishDiagram(new VisioDiagramPolishOptions {
    ResolveShapeOverlaps = true,
    MaximumConnectorLabelWidth = 1.6,
    FitHorizontalMargin = 0.6,
    FitVerticalMargin = 0.45
});
doc.Save();
```

`RelayoutAsGrid`, `RelayoutAsHorizontalStack`, and `RelayoutAsVerticalStack`
are deterministic and can reroute connectors whose endpoints are both inside a
page-backed selection.
`PolishDiagram` can also move crowded top-level shapes apart before it resolves
connector labels and fits the page, which is useful when a generated diagram has
reasonable structure but still needs a final visual cleanup pass.

## Learning from VSDX fixtures

OfficeIMO can inspect `.vsdx` files to learn which supported masters are present,
then generate OfficeIMO-owned masters from code. The source file is a learning
fixture, not a runtime template, so generated documents stay dependency-light
and deterministic.

```csharp
using OfficeIMO.Visio;

var doc = VisioDocument.Create("typed-shapes.vsdx");
doc.UseMastersByDefault = true;
doc.LearnMastersFromVsdx("DrawingWithShapes.vsdx",
    new[] { "Rectangle", "Ellipse", "Diamond", "Dynamic connector" });

var page = doc.AddPage("Page-1");
page.AddRectangle(2, 2, 2, 1, "Generated from OfficeIMO");
doc.Save();
```

## Optional native Visio validation

On Windows machines with Microsoft Visio installed, you can add a stronger
desktop compatibility gate without adding a compile-time Visio dependency.
OfficeIMO uses late-bound COM automation only when you call this helper.

```csharp
using OfficeIMO.Visio;

VisioDesktopValidationOptions options = VisioDesktopValidationOptions.RoundTripWithSvg();
options.SaveCopyPath = "diagram.visio-roundtrip.vsdx";
options.ExportDirectory = "visio-proof";

VisioDesktopValidationResult result = VisioDesktopValidator.Validate("diagram.vsdx", options);
if (result.IsAvailable && !result.IsValid) {
    throw new InvalidOperationException(string.Join(Environment.NewLine, result.Issues));
}
```

See `OfficeIMO.Examples/Visio/*` for more.

## Feature Scope (early)

- 📄 Pages: ✅ add/remove pages
- 🧱 Shapes: ✅ basic shapes from masters (rectangle, etc.), ✅ set text
- 🔗 Connectors: ✅ basic connectors between shapes
- 🧭 Diagram builders: ✅ flowchart builder with vertical and two-column continuation layouts plus branch routing, ✅ block diagram builder with grid regions and data/control flows, ✅ architecture builder with infrastructure components, regions, and routed data/control/dependency flows, ✅ network builder with zones, devices, links, and legends, ✅ swimlane builder with lanes, phases, activities, handoffs, and exception paths, ✅ org chart builder with hierarchy, assistants, team bands, vacancies, and external roles, ✅ timeline builder with date-scaled milestones and span lanes
- 🧰 Native stencils: ✅ built-in searchable catalogs for basic, flowchart, block-diagram, architecture, network, swimlane, org-chart, and timeline shapes
- 🎨 Style themes: ✅ reusable shape/connector/text styles and Modern/Office/Fluent/Technical/Minimal/Dark/Print authoring presets
- 🔎 Rich editing: ✅ recursive shape queries, shape/data/text/master/layer/hyperlink selectors, connector neighbor queries, page layers, shape and connector hyperlinks, bulk style/data/layer/hyperlink edits, align/distribute, resize-to-text, center content, and fit-to-content
- 🧩 VSDX learning fixtures: ✅ inspect supported masters without treating sample files as runtime templates
- 🧪 Validation: ✅ package/in-memory validators, ✅ optional Microsoft Visio desktop open, save-copy, and SVG/PNG/PDF export checks via late-bound COM

## Authoring units

Pages now remember a `DefaultUnit` (inches by default). When you create a page
with centimeters or millimeters, shape-adding overloads use that unit implicitly,
and the fluent shape builders follow that page unit as well, so you don't need helper conversions:

```
var page = doc.AddPage("A4 landscape", 29.7, 21.0, VisioMeasurementUnit.Centimeters);
page.AddRectangle(4.0, 15.0, 4.0, 2.5, "Rectangle"); // all values in cm
page.AddCircle(16.0, 15.0, 3.5, "Circle");           // diameter in cm
```

If you prefer, you can still pass an explicit unit:

```
page.AddRectangle(1.5, 1.0, 2.0, 1.0, "Rect", VisioMeasurementUnit.Inches);
```

## Connection points

You no longer need to add side connection points manually. The connector API
ensures side glue automatically when you specify `VisioSide.Left/Right/Top/Bottom`.
The old ensure method has been internalized.
- 🎨 Themes: ⚠️ minimal/default theme usage

This package is intentionally minimal at this stage and will expand over time.

## At a glance

- Create/Load/Save .vsdx (OPC packaging)
- Add simple pages, shapes, and connectors
- Fluent builder: `Page(...)`, `Rect(...)`, `Square(...)`, `Ellipse(...)`, `Circle(...)`, `Diamond(...)`, `Triangle(...)`, `Connect(...)`

## Why OfficeIMO.Visio (early)

- Minimal, no‑frills VSDX generation and reading using OPC + LINQ to XML
- Practical starting point for simple diagrams (pages, basic shapes, connectors)
- Designed to evolve as core scenarios are validated
