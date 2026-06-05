# OfficeIMO.Visio Assessment

Date: 2026-06-05
Branch/worktree: `codex/visio-market-review-20260605` at `C:\Support\GitHub\OfficeIMO-visio-market-review-20260605`

## Current Status Update

The Visio work has moved well beyond the first external-stencil slice. The current `origin/master` baseline includes merged graph/stencil work, native SVG/PNG preview export, premium gallery baselines, inspection snapshots, stencil profiles, data graphics, and broader diagram builders. Focused validation on this worktree passed:

```powershell
dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --framework net8.0 --filter "FullyQualifiedName~Visio"
```

Result: `738/738` Visio-filtered tests passed.

The current product slice includes:

- Package-backed stencil catalogs for installed Visio and external `.vssx` / `.vstx` / `.vsdx` sources.
- External multi-file stencil-pack discovery and examples, including repository-style packs such as Microsoft Integration/Azure stencil collections.
- A generic graph diagram builder with layered, grid, and radial layouts; cycles; disconnected components; zones; connector kinds; labels; hyperlinks; shape data on nodes; and shape data on connectors.
- Catalog selection helpers such as `FindBest`, `TryFindBest`, and graph builder overloads that let callers choose stencil nodes by catalog query instead of manually plucking masters.
- Robust gallery and showcase handling for duplicate generated IDs, connector ID reservation, metric-page sizing, optional pack probing, and stencil caption page fitting.
- Better visual behavior for package-backed stencil nodes: imported master artwork stays clean and captions are rendered as separate Visio text boxes when needed.
- Dependency-free native SVG and PNG preview export for OfficeIMO-authored diagrams, backed by premium visual baselines and structural/stencil profile snapshots.
- Versioned showcase summary evidence with top-level `proofTotals` and `evidenceTotals`, per-diagram preview/proof/evidence/visual-quality records, a browsable Evidence Coverage gallery section, and CI validation that can require native SVG/PNG previews plus inspection, stencil-profile, visual-quality, complete structural, and complete review proof for every diagram while exposing clean-quality and warning/error rollups.
- Shape Data data graphics that render badge/bar adornments as real local geometry, now with data-bar range validation to prevent invalid generated output.
- Typed fluent page targeting and selection editing for loaded diagrams, so create-or-edit workflows can use `ExistingPage(...)`, `PageOrAdd(...)`, and bulk shape/connector selection helpers instead of dropping to lower-level loops.
- Advanced fluent loaded-diagram selections for contained/intersecting geometry, connected components, shortest paths, User cells, hyperlinks, protection, connector layers, connector hyperlinks, and connector neighborhoods.
- Fluent replace-master and stencil-standardization helpers for loaded pages, backed by the existing page/selection replacement engine.
- Typed stencil migration maps for loaded documents, pages, and page-backed selections, so callers can standardize whole shape families by current stencil id, master name, shape `NameU`, or typed predicate in one first-match-wins pass.
- Catalog-query migration helpers and a basic-flowchart migration preset, so loaded basic/generated flowchart shapes can be upgraded to first-party or package-backed semantic stencils with less boilerplate.
- Non-mutating migration planning through `PlanStencilMigration(...)`, with stable `ToText()` reports plus `SaveText(...)`, `LoadText(...)`, and `FromText(...)` plan artifacts for reviewing loaded-diagram standardization before applying changes across process boundaries.
- Reviewed migration plans can be applied through `ApplyStencilMigration(plan, map)`, which validates page identity, shape identity, original text/master/stencil metadata, the matching rule, and replacement stencil before it mutates the loaded diagram.
- Network/infrastructure, architecture, org-chart, timeline, sequence, swimlane/process-map, cloud infrastructure, security/identity, Kubernetes/container, data/platform, and collaboration/business-process migration presets that use conservative label/name cues to upgrade common unstenciled legacy diagrams to semantic catalog stencils while skipping shapes already carrying OfficeIMO stencil metadata.
- Typed native container metadata/style editing for load-edit-save workflows, including `VisioContainerInfo`, option snapshots from loaded containers, native User-cell updates for margin/resize/lock/no-highlight/no-ribbon/style ids, reusable `VisioShapeStyle` application, fluent ID-based configuration, and metric-page refit that converts stored inch-based margins back to page units before resizing.
- Typed swimlane maintenance for load-edit-save workflows: generated swimlane diagrams now persist lane/phase/activity placement User cells, loaded pages can discover swimlane lanes/phases/activities, move activities between lane/phase cells, restack cells, and reroute affected connectors through page extensions or fluent loaded-page verbs.
- Fluent loaded-page relayout helpers now expose deterministic shape-id, connected-component, and Visio-native container-member relayout workflows over the typed selection engine, and the selection engine rejects non-finite spacing before it can generate invalid coordinates.

That changes the roadmap. The next bottleneck is no longer "can we load external stencils?" It is now "can we make generated diagrams look consistently premium and make real stencil packs easy to use at scale?"

For the live staged plan, see [officeimo.visio.roadmap.md](./officeimo.visio.roadmap.md).

## Executive Verdict

OfficeIMO.Visio is already past the toy stage. It can create, load, edit, preserve, validate, and round-trip real `.vsdx` packages. The current foundation includes pages, shape primitives, semantic flowchart aliases, masters, connection points, connector glue, groups, shape data, metadata, themes, package validation, and a small fluent API.

The product is not yet the "only Visio library a C# developer needs" because the high-value layer is still missing: automatic diagram construction, attractive defaults, reusable stencil catalogs, visual themes, layout engines, diagram-specific builders, import/export verification, and guided editing APIs. Today the user still thinks in coordinates and Visio XML-adjacent concepts. To win, OfficeIMO.Visio should let users think in diagrams: flowcharts, networks, org charts, swimlanes, process maps, Azure diagrams, containers, annotations, routes, styles, and data.

The strongest path is not to replace the existing model. Keep the current VSDX package engine. Add a first-party diagram authoring layer above it, and reuse `OfficeIMO.Drawing` only for portable visual intent: colors, text measurement, vector descriptors, paths, gradients, shadows, transforms, clipping, and simple drawing scenes. Visio still needs its own masters, ShapeSheet, connectors, layers, pages, stencils, and layout semantics.

## Current Inventory

### Project Shape

- `OfficeIMO.Visio` targets `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
- It references `OfficeIMO.Drawing`.
- It references `System.IO.Packaging` across target frameworks and `Microsoft.Bcl.AsyncInterfaces` for `net472`.
- The package metadata declares MIT and `OfficeIMO.Visio/LICENSE.MD` now matches the repository MIT license story.
- README now describes the broader Visio surface instead of calling the package early/minimal, but website and product docs still need an API-accuracy pass before public positioning.

### Authoring API

Implemented today:

- `VisioDocument.Create`, `Load`, `Save`, stream load/save.
- Multiple pages with size, default units, scale settings, grid/snap metadata.
- Basic and semantic shapes: rectangle, square, circle, ellipse, diamond, triangle, process, decision, data, preparation, manual operation, off-page reference, parallelogram, hexagon, trapezoid, pentagon.
- Shape styling: fill color, line color, line weight, line/fill pattern.
- Shape text, connector label text styling, and shape data.
- Connection points, automatic side glue, reconnect/retarget APIs.
- Connectors: dynamic, straight, right-angle, curved; start/end arrows; labels; line and label text styling; deterministic explicit waypoints that save and load back into the editing model, page-level placement/routing/line-jump direction policy, shape/connector-level route appearance, line-jump, reroute, placement, permeability, and splitting overrides, layout-grid sizing, and spacing cells, simple orthogonal route generation, branch/loop escape lanes in the flowchart builder, and first-class connector label placement.
- Dependency-free visual quality analysis for generated/loaded diagrams: page bounds, shape overlap, explicit connector/shape crossings, connector label bounds, connector label/shape overlaps, connector label/label overlaps, deterministic connector label cleanup, high-level diagram polish passes, and a generated gallery runner that can attach optional Microsoft Visio desktop round-trip/export proof.
- Deterministic text measurement for shape text boxes and connector label boxes through `OfficeIMO.Drawing`.
- Connector-aware page fitting and centering that include explicit waypoints and connector label boxes.
- One-call page/document polish helpers that can resize connector labels, resolve connector label collisions, optionally resize shapes to text, and fit pages to content.
- Groups and nested shapes, reparenting, ungrouping, child hierarchy preservation.
- Masters: built-in generated masters, document registry, `.vsdx`-based structural learning for supported names, master-backed page instances.
- Fluent API for document/page/shape/connector basics, including loaded-page targeting, typed selection edits, duplicate-shape workflows with semantic copy IDs, topology-aware updates, connector-neighborhood updates, replace-master workflows, reusable stencil migration maps, catalog-query migration rules, domain migration presets, and non-mutating migration reports.
- Reusable authoring style themes: `VisioShapeStyle`, `VisioConnectorStyle`, text-aware styles, and `VisioStyleTheme` with Modern, Office, Fluent, Technical, Minimal, Dark, and Print presets. Flowchart/block/architecture builders and selection/fluent APIs can consume the same style theme.

Current weakness:

- The fluent API is friendly and can now target loaded pages, bulk-edit typed selections, edit by geometry/path/component/connectors, standardize masters, run typed stencil migration maps, use catalog-query/domain preset migration helpers, plan/report/persist migrations, and replay approved migration plans with drift validation, but it is still coordinate-driven for new shape placement.
- Styling still exposes some Visio-index details for line/fill patterns, but reusable style objects and presets now give high-level code a safer default path.
- There is now a first-pass diagram DSL for flowcharts, block diagrams, dependency diagrams, architecture diagrams, network diagrams, swimlane process maps, org charts, and date-scaled timelines, plus reusable data-driven gallery scenarios for security, data, network, and process governance. The remaining gap is breadth and depth: richer branch layout, joins, advanced lane assignment, richer org chart variants, deeper BPMN-style process semantics, and automatic obstacle-aware routing.
- There is now first-pass deterministic connector routing, label placement, label cleanup, visual quality analysis, and a CI-friendly `EnsureVisualQuality(...)` quality gate for explicit/manual routes, but no full diagram-level routing engine with global obstacle avoidance, line crossing minimization, or multi-pass label optimization.
- Premium enterprise, technical, cloud, process, print-safe, and dark-safe presets exist and are used in the gallery. The remaining issue is art direction and breadth: more scenarios need to reach the same reviewed bar before the package can claim market-leading output quality.

### Editing And Round-Trip Fidelity

Implemented today:

- Load parses pages, shapes, connectors, masters, theme, metadata, style sheets, page sheets, ShapeSheet cells, shape text, rich text preservation, connection rows, custom geometry, connector layout/control sections, and many preserved XML fragments.
- The tests show strong attention to preserving unknown or unsupported content.
- Existing edit APIs cover text, data, style, grouping, connector reconnection, and connector retargeting.

Current weakness:

- Initial query/selection/layout helpers now cover recursive shapes, id/name/master/data/text/layer/hyperlink/User-cell/typed-Shape-Data lookups, page layers, first-pass native containers, typed container metadata/style updates, native comments, semantic callouts/annotations, connector neighbors, bounds, fit-to-content, center content, align/distribute, resize-to-text, connector routing/labels, same-page selection duplication with semantic copy IDs, swimlane lane/phase/activity discovery and activity move/restack workflows, deterministic shape/container-member relayout, and bulk style/data/layer/hyperlink/User-cell/Shape-Data/geometry edits. The fluent layer now exposes loaded-page targeting, common typed shape/connector selection edits, id-based geometry/path/component selectors, connector-neighborhood edits, duplicate shape/selection workflows, native container add/remove/refit/relayout/configure/style workflows, native page/shape comment add/update/resolve/reopen/remove workflows, swimlane move/relayout workflows, replace-master workflows, reusable stencil migration maps, catalog-query migration helpers, basic-flowchart/network-infrastructure/architecture/org-chart/timeline/sequence/swimlane/cloud/security/Kubernetes/data/collaboration migration presets, dry-run migration reports, persisted migration-plan artifacts, and validated approved-plan application for load-edit-save workflows. The remaining gap is advanced diagram-level editing: deeper swimlane metadata/auto-assignment, richer comment threading/author workflows, richer nested/container behavior, advanced resize-to-content, and broader diagram-level relayout/polish workflows.
- Advanced Visio concepts are only partly first-class: page layers, reusable background pages, page print/setup/lock settings, page-level placement/routing/line-jump direction/layout-grid/spacing settings, shape/connector-level placement/routing/route-appearance/line-jump/reroute/permeability/splitting overrides, shape/connector hyperlinks, generic User cells, typed Shape Data rows, first-pass native containers, native comments, semantic callouts/annotations, and shape/connector protection now round-trip as typed objects, but data graphics, legends, formulas beyond targeted cells, and many ShapeSheet sections beyond preservation still need typed APIs.
- Load fidelity is strong for preservation but not yet a complete typed object model.

### Validation And Tests

Evidence gathered in this assessment:

- Visio-related test declarations: 717 facts/theories by file scan.
- Latest focused Visio test run passed `738/738`:

```powershell
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --framework net8.0 --filter "FullyQualifiedName~Visio"
```

Notes:

- The build emitted existing nullable warnings from unrelated test files.
- The test slice now includes optional native Microsoft Visio application open, save-copy, and SVG export checks when Visio is installed. It still needs richer repair diagnostics and visual screenshot/export baseline comparison.

### Assets And Samples

Existing assets:

- `Assets/VisioTemplates/DrawingEmpty.vsdx`
- `Assets/VisioTemplates/DrawingWithRectangle.vsdx`
- `Assets/VisioTemplates/DrawingWithShapes.vsdx`
- `Assets/VisioTemplates/DrawingWithSomeShapes.vsdx`
- `Assets/VisioTemplates/DrawingWithLotsOfShapresAndArrows.vsdx`
- `Assets/VisioTemplates/DrawingWithInfoAndShapes.vsdx`
- `Assets/VisioTemplates/DrawingWithJenkinsDiagram.vsdx`

Important boundary: these `.vsdx` assets should be treated as learning fixtures and golden references, not as runtime templates that OfficeIMO depends on. Their main value is to show how Visio-authored packages, masters, geometry, page parts, connectors, styles, and ShapeSheet fragments are structured, and to compare OfficeIMO-generated output against real Visio output. We should learn from them, generate our own clean model, and avoid designing the library around ingesting whole templates as the normal authoring path.

Existing examples:

- Basic document.
- Fluent basic diagram.
- Rectangle connections.
- Connection points.
- Shape catalog.
- Typed/master-based shape catalog.
- Asset catalog and master extraction/import.
- Read existing Visio document.

Current weakness:

- The examples are useful API smoke tests, not a polished gallery.
- There are now copyable builder examples for flowchart, block diagram, cloud/service architecture, network map, swimlane process diagrams, org charts, and timeline roadmaps, with gallery-backed data-driven examples for rack/server operations and process governance. Remaining target examples include deeper BPMN-style process detail, dependency graphs, and more polished visual baselines.

## Reuse From OfficeIMO.Drawing

The in-progress `OfficeIMO.Drawing` work in the main checkout is directly useful, but as a shared visual intent layer rather than a Visio document replacement.

Reuse strongly:

- `OfficeColor` as the common color model already used by Visio.
- `OfficeTextMeasurer`, `OfficeFontInfo`, and `OfficeTextMetrics` for text box sizing, label wrapping, node auto-sizing, and layout estimation.
- `OfficeShape`, `OfficeDrawing`, `OfficeDrawingShape`, `OfficePoint`, `OfficePathCommand`, and path/polygon descriptors for freeform custom masters and small reusable visual assets.
- `OfficeLinearGradient`, `OfficeShadow`, `OfficeTransform`, and `OfficeClipPath` as cross-format visual intent that Visio can map to native cells/sections where supported.
- `OfficeImageReader` and `OfficeImageFit` for images/icons embedded in Visio shapes or master artwork.

Do not reuse as the core:

- `OfficeDrawing` is a local top-left canvas. Visio is a page/ShapeSheet/master/connector system with page coordinates, shape local coordinates, glue, routing, pages, stencils, layers, and formulas.
- Visio needs semantic diagram objects and VSDX package fidelity that a generic drawing scene should not own.

Recommended adapter shape:

```csharp
namespace OfficeIMO.Visio.Drawing;

public static class VisioDrawingExtensions {
    public static VisioShape AddDrawing(this VisioPage page, OfficeDrawing drawing, double x, double y, VisioDrawingOptions? options = null);
    public static VisioMaster RegisterDrawingMaster(this VisioDocument document, string nameU, OfficeDrawing drawing, VisioDrawingMasterOptions? options = null);
}
```

This keeps the boundary clean: Drawing describes reusable visual primitives; Visio maps them into masters/groups/shapes.

## Competitive Position

Microsoft documents `.vsdx` as OPC plus XML, which validates OfficeIMO's current dependency-free, no-COM direction. The commercial competitor to beat remains Aspose.Diagram, currently published as a broad paid Visio manipulation/conversion API with support for many Visio formats and exports. There is also market movement toward open-source Visio tooling, including Aspose's own public "FOSS coming soon" positioning.

The opportunity is clear: a dependency-light, MIT, C#-native, server-safe, cross-platform Visio library with excellent authoring ergonomics would occupy a real gap. OfficeIMO can win if it becomes easier than raw XML/Interop and more acceptable than commercial black-box libraries.

Sources checked:

- Microsoft Visio `.vsdx` file format overview: https://learn.microsoft.com/en-us/office/client-developer/visio/introduction-to-the-visio-file-formatvsdx
- Microsoft programmatic manipulation guidance: https://learn.microsoft.com/en-us/office/client-developer/visio/how-to-manipulate-the-visio-file-format-programmatically
- Aspose.Diagram NuGet package: https://www.nuget.org/packages/Aspose.Diagram/
- Aspose.Diagram documentation: https://docs.aspose.com/diagram/net/
- Aspose.Diagram FOSS positioning: https://products.aspose.org/diagram/

## What Complete Should Mean

OfficeIMO.Visio should define "complete" in layers, not as one giant ShapeSheet clone.

### Level 1: Reliable VSDX Core

Goal: every file OfficeIMO creates opens cleanly in Visio, validates structurally, and round-trips without destroying unsupported content.

Needed:

- License cleanup.
- Explicit dependency policy: either accept `System.IO.Packaging` as infrastructure or replace it with an internal OPC/ZIP writer for strict zero external NuGet dependencies.
- Native Visio app validation harness for generated examples.
- Golden package comparison against Visio-authored assets, with those assets used as structural references rather than required runtime templates.
- Package preflight with actionable diagnostics.
- Compatibility matrix: Windows Visio, Visio web, LibreOffice/draw.io where relevant, and OfficeIMO load/save.

### Level 2: Friendly Basic Authoring

Goal: users can create clean simple diagrams without knowing Visio internals.

Needed:

- Diagram themes: Office, Fluent, Modern, Minimal, Technical, Dark, Print. Initial reusable presets exist and now carry text styles so dark and saturated themes remain readable.
- Shape style objects instead of raw pattern indexes everywhere. Initial `VisioShapeStyle` and `VisioConnectorStyle` exist, with raw pattern values still exposed for advanced control.
- Text style objects: initial whole-shape text style support exists for font family, color, size, bold, italic, underline, horizontal/vertical alignment, margins, text-block transform cells, and text background; rich text runs, wrap, and auto-fit remain.
- Connector style objects: arrow, elbow/orthogonal/curve, dashed, label placement.
- Page helpers: portrait/landscape presets, margin, fit-to-content, center content.
- Built-in polished examples for basic flowchart, block diagram, network, cloud architecture, and swimlane. Initial generated gallery coverage exists for all five.

### Level 3: Layout And Diagram DSL

Goal: users describe relationships; OfficeIMO places shapes and routes connectors.

Needed:

- `FlowchartBuilder`: steps, decisions, branches, joins, off-page references. Initial builder exists with branch escape-lane routing.
- `BlockDiagramBuilder`: blocks, buses, dashed control flow, data flow, labels. Initial builder exists with grid regions and semantic data/control links.
- `DependencyDiagramBuilder`: nodes and directed dependencies with automatic deterministic layered DAG layout. Initial builder exists for acyclic dependency maps.
- `ArchitectureDiagramBuilder`: infrastructure components, regions, data/control/dependency flows. Initial builder exists for dependency-free cloud/service diagrams.
- `NetworkDiagramBuilder`: nodes, groups, containers, zones, edges, legends. Initial builder exists with zones, typed devices, routed links, and legends. Initial `NetworkTopologyDiagramBuilder` also exists for graph-first automatic topology layout with mesh/cycle support and subnet/background zones derived from automatically placed nodes.
- `SwimlaneDiagramBuilder`: lanes, phases, activities, handoffs, exception paths, deterministic routed flows, and same-cell activity stacking. Initial builder exists with editable lane/phase scaffold shapes and semantic activity placement.
- `OrgChartDiagramBuilder`: hierarchy, assistant nodes, team bands, vacancies, external roles, and routed reporting lines. Initial builder exists with deterministic top-down layout.
- Layout algorithms: layered DAG, tree, grid, force/relax for simple network maps, swimlane lane assignment, orthogonal routing, collision avoidance, label avoidance.
- Deterministic layout so CI baselines are stable.

### Level 4: Stencils And Masters

Goal: native stencil support is not a side feature; it becomes the main scale lever.

Needed:

- Read `.vssx` and `.vstx` as intentional stencil/master sources. Initial package-backed catalog loader reads `.vsdx`/`.vssx`/`.vstx` master metadata and exposes supported generated masters without runtime template dependency.
- Use `.vsdx` samples primarily as structural references and compatibility baselines; avoid depending on `.vsdx` drawings as reusable runtime templates.
- Preserve and reuse imported stencil masters where licensing and package structure are clear, not only supported generated equivalents.
- Create custom stencils from code. Initial fluent generated catalog builder and dependency-free OfficeIMO XML manifest save/load exist for code-defined reusable palettes.
- Ship first-party generated stencil packs with safe licensing:
  - Basic Shapes (initial generated catalog added)
  - Flowchart (initial generated catalog added)
  - Architecture (initial generated catalog added)
  - Network (initial generated catalog added)
  - Swimlane (initial generated catalog added)
  - Servers and Devices
  - Cloud Generic
  - Azure-like generic pack, carefully avoiding trademarked copied artwork unless licensing is clear
  - Security and Identity
  - Containers/Kubernetes generic
- Master search by category, alias, tag, keyword, and icon preview metadata. Initial generated catalogs now expose ranked `Search(...)`, `InCategory(...)`, aliases, tags, category lists, and `IconNameU`.
- `UseStencil(...)`, `Shape("Azure.VirtualMachine")`, `Shape("Network.Switch")` style APIs. Initial page/fluent `AddStencilShape`/`Stencil` helpers exist for generated catalogs, including all-catalog string placement such as `AddStencilShape("net.switch", ...)`.

### Level 5: Rich Editing

Goal: OfficeIMO can modify existing diagrams predictably.

Needed:

- Query API:
  - by id/name/master/type/layer/data/text (initial id/name/master/data/text support added)
  - incoming/outgoing connectors (initial support added)
  - group descendants (initial support added)
  - page bounds and intersections (initial bounds support added; intersections still needed)
- Bulk operations:
  - restyle selected shapes (initial shape/connector selection styling added)
  - replace master (initial page/selection/stencil replace-master editing API added)
  - align/distribute (initial support added)
  - resize to text (initial deterministic text measurement support added)
  - deterministic connector route cleanup (initial explicit waypoint/orthogonal routing and label placement support added)
  - relayout selection (initial deterministic grid/horizontal/vertical selection relayout with internal connector rerouting added)
  - duplicate page/selection (initial full-page duplication, optional copied background-page dependency, same-page selection duplication with internal connector remapping, typed duplication options, and fluent semantic copy IDs added)
  - semantic callouts/annotations with leader connectors (initial support added)
  - native container membership editing (initial add/remove/refit APIs and fluent ID-based loaded-page workflows added)
  - native comments (typed page/shape comments with `/visio/comments.xml` save/load plus update/resolve/reopen/remove and fluent loaded-page workflows added)
- First-class layers, hyperlinks, User cells, typed Shape Data, first-pass containers, native comments, and semantic callouts/annotations are now started; continue with richer container metadata/styles, swimlanes, richer comment threading/author workflows, and data graphics.
- Typed shape data with labels, prompts, types, and formats is now started; continue with richer formulas, data linking, and schema-level helpers.
- Safe unsupported-content policy: preserve by default, expose raw XML escape hatches when needed.

### Level 6: Visual Verification And Export

Goal: the library can prove diagrams are visually good, not just XML-valid.

Needed:

- Native Visio COM validation on Windows agents or local optional test suite:
  - open generated VSDX (initial optional late-bound helper added as `VisioDesktopValidator`)
  - save-as VSDX (implemented as optional save-copy validation)
  - export SVG/PNG/PDF (implemented as optional export validation; SVG covered by test)
  - collect repair/open errors
- Generated gallery sample runner and dependency-free quality analyzer are now available as a first verification layer; visual baselines for gallery diagrams are still needed.
- Optional OfficeIMO-native SVG export for CI and web documentation.
- Reuse `OfficeIMO.Pdf` and `OfficeIMO.Drawing` where it makes sense for PDF/SVG previews, but keep VSDX as source of truth.

## Recommended Public API Direction

The current low-level API should stay, but most users should enter through builders.

Example target shape:

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

VisioDocument.Create("property-buying-flowchart.vsdx")
    .Flowchart("Property buying Flowchart", flow => flow
        .Theme(VisioStyleTheme.Modern())
        .Title()
        .Layout(VisioFlowchartLayout.TwoColumnContinuation)
        .Start("agent", "Start with an agent\nyou trust")
        .Step("consult", "Consult with agent to\ndetermine your property\nwants and needs")
        .Step("paperwork", "Review and complete\npaperwork")
        .Step("loan", "Go to preferred lender,\nget pre-qualified and\npre-approval for loan\namount")
        .Step("market", "With agent, analyze\nmarket to choose\nproperties of interest")
        .Step("view", "View properties\nwith agent")
        .OffPage("jump", "A")
        .Continue("resume", "A")
        .Step("offer", "Select ideal property\nand write offer to\npurchase")
        .Decision("negotiate", "Negotiate\n& Counteroffer:\nAgreement?")
        .Branch("negotiate", "No", "market")
        .Branch("negotiate", "Yes", "contract")
        .Step("contract", "Accept the contract")
        .Step("underwriting", "Secure underwriting,\nobtain loan approval")
        .Step("closing", "Select/Contact closing\nattorney for title exam\nand title insurance")
        .Step("inspection", "Schedule inspection\nand survey")
        .End("close", "Close on the\nproperty"))
    .Save();
```

The key is that the builder owns layout, continuation markers, connectors, text wrapping, colors, and defaults. Users can still override every shape when they need to.

## Priority Roadmap

### P0: Make The Foundation Product-Ready

- Keep the Visio license story clean: package metadata and local license text now match MIT, and future changes should not reintroduce conflicting local license wording.
- Decide strict dependency-free policy:
  - practical: keep `System.IO.Packaging`;
  - strict: replace with internal OPC writer/reader over `ZipArchive`.
- Update README to reflect actual capabilities and limitations.
- Add local optional Visio validation harness. Initial open/save-copy/export helper exists and gallery runs can now attach optional desktop round-trip/export proof; expand it to repair diagnostics and visual baseline capture.
- Add generated gallery sample runner for beautiful reference diagrams. Initial runner exists with package validation, visual-quality validation, machine-readable showcase artifact summaries, deterministic inspection/profile/visual-quality proof artifacts, explicit preview/proof evidence totals, browsable showcase gallery output, and optional desktop proof result data.
- Add polished theme/style model.

### P1: First "Wow" Authoring Layer

- Build `FlowchartBuilder` end to end, including richer auto layout, branch routing, off-page continuation, text wrapping, and attractive default themes.
- Build `BlockDiagramBuilder` for data/control-flow diagrams with solid/dashed connectors and 2.5D block style if desired.
- Add richer `FitToContent`, `Align`, `Distribute`, routing, and connector label placement.
- Add `OfficeDrawing` to Visio master/group adapter.
- Add visual baseline tests for the gallery outputs.

### P2: Stencils As A Platform

- Add `.vssx` and `.vstx` master catalog loading. Initial package-backed `VisioStencilPackageCatalog.Load(...)` exists for `.vsdx`/`.vssx`/`.vstx`, defaulting to supported generated masters and optional unsupported generic entries.
- Keep `.vsdx` sample files as learning/verification fixtures, not as the recommended authoring template mechanism.
- Preserve imported unsupported masters as package parts where possible.
- Add first-party generated stencil packs. Initial Basic/Flowchart/Block Diagram/Architecture/Network/Swimlane/Org Chart/Timeline generated catalogs exist.
- Add stencil search, aliases, categories, and fluent shape factories. Initial ranked search, category filtering, alias/tag metadata, icon metadata, package-backed catalog loading, and page/fluent all-catalog placement exist for generated catalogs.
- Add custom stencil export. Initial custom generated catalog authoring plus OfficeIMO-native XML manifest save/load exists; export to real reusable `.vssx`/package format is still needed.

### P3: Editing And Import Parity

- Add query/selection APIs.
- Continue layer/hyperlink/container/comment depth and add richer swimlane/callout support.
- Continue typed shape-data depth with schema-level helpers, formulas, and data-linking metadata.
- Add replace-master and relayout-selection operations.
- Expand round-trip tests using Visio-authored files.

### P4: Export And Visual QA

- SVG preview/export.
- PDF preview/export via OfficeIMO.Pdf where feasible.
- Native Visio COM export harness for Windows verification.
- CI-friendly visual diff artifacts for gallery diagrams. Initial issue-counting report and throwing quality gate exist; pixel/baseline diff artifacts are still needed.

## Design Rules To Avoid Dead Ends

- Keep `OfficeIMO.Visio` as the VSDX authority; do not make `OfficeIMO.Drawing` own Visio semantics.
- Keep unknown XML preservation as a first-class design rule.
- Prefer semantic builders over more raw overloads.
- Make layouts deterministic.
- Keep the public API dependency-free and server-safe; no Microsoft Office automation in the core library.
- Treat local Visio automation as verification tooling only.
- Treat bundled `.vsdx` drawings as structure-learning and regression fixtures only.
- Separate generated first-party stencils from imported third-party stencils.
- Keep style/theme objects portable enough to share with PDF/PowerPoint later, but map them format-specifically.

## Bottom Line

OfficeIMO.Visio has a surprisingly solid package and round-trip core. The next leap is product design: diagram builders, layout, stencils, themes, gallery-quality output, and verification. If those layers are added without disturbing the current preservation-focused VSDX engine, OfficeIMO.Visio can plausibly become the best open C# option for dependency-light Visio creation and editing.
