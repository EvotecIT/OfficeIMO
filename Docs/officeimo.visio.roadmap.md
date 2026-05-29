# OfficeIMO.Visio Roadmap

Date: 2026-05-28
Branch/worktree: `codex/visio-external-stencil-packs` at `C:\Support\GitHub\OfficeIMO-visio-external-stencil-packs`
Active PR: https://github.com/EvotecIT/OfficeIMO/pull/1865

## Where We Are

OfficeIMO.Visio is no longer just a basic VSDX writer. The current branch has a usable, dependency-light Visio authoring stack:

- VSDX package creation, loading, editing, saving, validation, and preservation of unsupported content.
- Pages, shapes, connectors, connection points, groups, layers, hyperlinks, User cells, typed Shape Data, protection, page settings, backgrounds, metadata, themes, style sheets, and master-backed page instances.
- Fluent page and document authoring for lower-level diagrams.
- High-level builders for flowcharts, block diagrams, dependency diagrams, architecture diagrams, networks, network topology, swimlanes, org charts, timelines, sequences, and generic graphs.
- Reusable style themes and local node/edge style overrides.
- Connector routing, label placement, label cleanup, page fitting, deterministic text measurement through `OfficeIMO.Drawing`, and visual quality analysis.
- First-party generated stencil catalogs and package-backed stencil catalogs.
- External `.vssx`, `.vstx`, and `.vsdx` catalog loading, including multi-package external stencil repositories.
- Native/installed Visio stencil discovery when Visio is available, without making Office automation part of the core library.
- Optional desktop Visio validation/export helpers for local proof, separate from the dependency-free core.
- Generated showcase examples and preview output for visual inspection.

The branch is green at this checkpoint, but it should be treated as the current staged PR, not the final Visio destination.

## Product Definition

The target is not "write a few shapes into a VSDX." The target is:

> A C# developer can describe a real process, topology, architecture, org structure, or graph, pick a native or external stencil catalog, choose a theme, and get a Visio file that opens cleanly, edits naturally, carries useful data, and looks professional without hand-managing coordinates.

That implies three durable layers:

- **VSDX core:** package, ShapeSheet, masters, pages, preservation, validation, and round-trip fidelity.
- **Diagram intelligence:** builders, layout, routing, labels, quality gates, themes, containers, zones, and metadata.
- **Stencil platform:** first-party generated catalogs, native Visio catalogs, external packs, package-backed masters, search, gallery, and authoring/export.

`OfficeIMO.Drawing` remains useful as a shared visual-intent layer for colors, text metrics, paths, simple vector descriptions, and preview/export work. It should not own Visio semantics.

## Recently Completed

### Foundation

- Load/save/edit VSDX without requiring Visio.
- Optional native Visio desktop validation and export.
- Validator hardening for generated and loaded diagrams.
- Master-backed diagram generation without defaulting to generated geometry where imported masters should be preserved.
- Text style serialization fixes for native Visio compatibility.
- Ellipse/circle preview geometry corrections.

### Diagram Builders

- Flowchart, block, dependency, architecture, network, network topology, swimlane, org chart, timeline, sequence, and generic graph builders.
- Layered/grid/radial graph layouts.
- Basic cycle and disconnected-component support for generic graphs.
- Zones and background containers for graph/topology diagrams.
- Connector kinds, labels, hyperlinks, and shape data on graph edges.
- Node hyperlinks, Shape Data, and style overrides.

### Stencils

- First-party generated catalogs.
- Package-backed catalog loading from `.vssx`, `.vstx`, and `.vsdx`.
- Installed Visio stencil discovery.
- External multi-file stencil-pack examples.
- Catalog query helpers: `FindBest`, `TryFindBest`, category/tag/alias search, and graph node overloads.
- Stencil gallery robustness around generated IDs, connector IDs, metric-page sizing, and optional-pack probing.

### Examples And Showcase

- Basic generated examples.
- Graph examples with native and external stencil nodes.
- External stencil gallery examples.
- Microsoft Integration/Azure pack example path.
- Showcase preview generation with optional Visio export proof.

## Immediate P0

These are the items that still block "premium Visio library" positioning.

1. **Merge the current green PR.**
   Keep PR #1865 as the active external stencil/graph slice unless new review feedback appears.

2. **Fix the public docs to match the real API.**
   Website and product docs should use `VisioDocument`, diagram builders, and stencil catalogs, not older pseudo-APIs.

3. **Make examples visibly better.**
   The library now supports richer output, but the gallery needs stronger art direction:
   - fewer plain rectangles;
   - more external/native stencil usage;
   - better spacing, hierarchy, and caption treatment;
   - consistent typography;
   - restrained but polished themes;
   - examples that show actual business value, not only API coverage.

4. **Add a visual baseline workflow.**
   Keep XML validation, but add generated SVG/PNG baselines for the showcase so regressions are visible.

5. **Resolve the Visio license conflict.**
   The package metadata says MIT, while `OfficeIMO.Visio/LICENSE.MD` is restrictive. This must be corrected before public messaging claims a clean open-source package story.

## P1: Premium Diagram Experience

Goal: generated diagrams should be credible without manual post-editing.

- Add a richer theme catalog with tuned typography, margins, contrast, connector styling, and label rules.
- Add builder-level `Theme(...)` presets for enterprise, technical, cloud, process, print, and dark-safe diagrams.
- Add diagram-specific visual defaults:
  - architecture zones and trust boundaries;
  - network subnets, racks, and device groupings;
  - flowchart continuation and branch retry patterns;
  - swimlane phase/lane readability;
  - sequence activations and notes;
  - dependency graph critical-path highlighting.
- Add automatic label collision cleanup for dense graph diagrams.
- Improve orthogonal routing around zones and large node groups.
- Add deterministic "polish passes" that can be applied after any builder.

## P2: Stencil Platform

Goal: native and external stencil packs feel first-class, not like shortcuts.

- Preserve more imported master/package content for unsupported external masters.
- Improve master metadata extraction: categories, keywords, preview/icon metadata, dimensions, connection points, and aliases.
- Add package-backed master reuse without forcing generated fallback where a real master exists.
- Add first-party generated stencil packs for:
  - servers and devices;
  - cloud generic;
  - security and identity;
  - containers/Kubernetes generic;
  - data/platform services;
  - collaboration/business process symbols.
- Add a real stencil gallery document builder for catalog review and debugging.
- Add custom stencil export, ideally to reusable package-backed stencil form when feasible.
- Document native Visio stencil discovery paths and external pack usage patterns.

## P3: Real Graphs And Data-Driven Diagrams

Goal: users can feed real inventories, dependencies, or workflows into OfficeIMO.Visio.

- Add graph import helpers from simple node/edge records.
- Add stable ID and diff-friendly regeneration guidance.
- Add graph clustering/grouping APIs.
- Add dependency cycle presentation instead of only rejection where the diagram type allows cycles.
- Add legends based on used node/edge types.
- Add data-driven examples:
  - Azure/application dependency map;
  - Active Directory identity/authentication flow;
  - CI/CD pipeline and build-agent topology;
  - Kubernetes/service-mesh topology;
  - incident/runbook sequence;
  - network segmentation diagram.

## P4: Editing Existing Diagrams

Goal: OfficeIMO can safely update diagrams created elsewhere.

- Expand replace-master operations for loaded external masters.
- Add richer selection queries: geometry intersection, contained-in-zone, connected component, path search, and data predicates.
- Add comment APIs and richer annotation/callout editing.
- Add containers and swimlanes as deeper typed concepts.
- Add data graphics and richer Shape Data schema helpers.
- Add safe relayout of selected subsets while preserving unsupported content.
- Add round-trip tests against more Visio-authored assets.

## P5: Verification And Export

Goal: prove quality continuously.

- Add first-party SVG preview/export where feasible.
- Use `OfficeIMO.Pdf` and `OfficeIMO.Drawing` for optional previews where they fit, without making them the VSDX source of truth.
- Expand desktop Visio validation to collect repair dialogs, export failures, and visual artifacts.
- Add CI artifacts for generated showcase previews.
- Add visual-diff thresholds for premium examples.
- Track compatibility with Microsoft Visio desktop, Visio web, and common VSDX consumers where practical.

## API Shape To Prefer

Prefer this kind of usage:

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

var catalog = VisioStencilPackageCatalog.LoadMany(
    VisioStencilPackageCatalog.EnumeratePackageFiles(@"C:\Stencils\Azure"));

VisioDocument.Create("architecture.vsdx")
    .GraphDiagram("Production topology", graph => graph
        .Title()
        .Layout(VisioGraphLayout.Layered)
        .Direction(VisioGraphDirection.LeftToRight)
        .StencilNode("gateway", "API Gateway", catalog, "API Management")
        .StencilNode("events", "Events", catalog, "Event Grid")
        .Node("worker", "Worker")
        .Node("database", "Database", VisioGraphNodeKind.Data)
        .Zone("runtime", "Runtime", "gateway", "events", "worker")
        .ControlEdge("gateway-events", "gateway", "events", "publish")
        .DataEdge("worker-db", "worker", "database", "write")
        .EdgeShapeData("worker-db", "Protocol", "SQL", "Protocol", VisioShapeDataType.String))
    .Save();
```

Keep this as the north star: callers describe the graph and the catalog, OfficeIMO manages layout, page fitting, captions, connectors, labels, metadata, and Visio package correctness.

## Design Rules

- Keep the core dependency-light and server-safe.
- Do not require Visio for generation; use Visio only for optional validation/export.
- Treat `.vsdx` samples as learning and regression fixtures, not runtime templates.
- Support native and external `.vssx` packs directly.
- Preserve unknown XML by default.
- Prefer semantic builders over coordinate-heavy sample code.
- Make generated IDs deterministic but collision-safe.
- Keep layout deterministic for CI.
- Put reusable behavior in `OfficeIMO.Visio`; keep examples and website docs thin.
