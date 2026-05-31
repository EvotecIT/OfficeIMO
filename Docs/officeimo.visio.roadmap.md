# OfficeIMO.Visio Roadmap

Date: 2026-05-31
Current branch/worktree: `codex/visio-premium-roadmap` at `C:\Support\GitHub\OfficeIMO-visio-premium-roadmap`
Last major Visio PR merged: https://github.com/EvotecIT/OfficeIMO/pull/1865

## Where We Are

OfficeIMO.Visio is no longer just a basic VSDX writer. The current branch has a usable, dependency-light Visio authoring stack:

- VSDX package creation, loading, editing, saving, validation, and preservation of unsupported content.
- Pages, shapes, connectors, connection points, groups, layers, hyperlinks, User cells, typed Shape Data, protection, page settings, backgrounds, metadata, themes, style sheets, and master-backed page instances.
- Fluent page and document authoring for lower-level diagrams.
- High-level builders for flowcharts, block diagrams, dependency diagrams, architecture diagrams, networks, network topology, swimlanes, org charts, timelines, sequences, and generic graphs.
- Reusable style themes, premium enterprise/technical/cloud/process/print/dark-safe presets, and local node/edge style overrides.
- Connector routing, obstacle-aware routing around unrelated shapes, optional zone/container/group/adornment-aware routing, connector-crossing-aware route scoring, deterministic page-level routing optimization passes, multi-waypoint dogleg candidates for dense crossings, label placement, whole-page connector label optimization passes, zone-aware and connector-path-aware label cleanup, page fitting, deterministic text measurement through `OfficeIMO.Drawing`, and visual quality analysis.
- Header-style region/zone captions for architecture, block, network, topology, and graph builders, including layout clearance and quality-analyzer handling for generated caption adornments.
- First-party generated stencil catalogs and package-backed stencil catalogs.
- External `.vssx`, `.vstx`, and `.vsdx` catalog loading, including multi-package external stencil repositories.
- Native/installed Visio stencil discovery when Visio is available, without making Office automation part of the core library.
- Optional desktop Visio validation/export helpers for local proof, separate from the dependency-free core.
- Generated showcase examples and preview output for visual inspection.
- Premium showcase diagrams with validated VSDX packages plus local PNG/SVG preview proof.
- A reusable eight-diagram `VisioPremiumGallery` plus approved PNG/SVG baseline fixtures and PNG diff artifacts for Visio desktop preview regression.
- Deterministic inspection snapshots and structural diffs through `CreateInspectionSnapshot()` / `VisioInspectionDiff`, including connector label placement coordinates for label-layout regression proof.
- Deterministic stencil usage profiles through `CreateStencilProfile()`, including generated-master, package-backed, basic-geometry, stencil-backed, Shape Data key, semantic-kind, catalog, category, keyword, alias, tag, icon identity, package preview-image content type/extension, learned native connection points, source default dimensions/units, source-package, placed/source-dimension, and connection-point summaries that survive save/load for package-backed masters and generated stencil placements.

The external stencil and graph slice from PR #1865 is merged. The next checkpoint is no longer proving that the Visio core can generate useful diagrams; it is making the public package story clean, the showcase reproducible, and the generated output visually strong enough to support premium positioning.

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
- Graph, architecture, network, flowchart, block-diagram, swimlane, timeline, and sequence stencil nodes selected from a catalog preserve catalog provenance in stencil profiles, and built-in catalog stencil nodes inherit diagram theme styling unless they come from package-backed external artwork.

### Examples And Showcase

- Basic generated examples.
- Graph examples with native and external stencil nodes.
- External stencil gallery examples.
- Microsoft Integration/Azure pack example path.
- Showcase preview generation with optional Visio export proof.
- Premium gallery baseline proof through `VisioPremiumVisualBaselineTests`, including approved PNG/SVG fixtures, PNG pixel-diff artifacts on failure, tolerance knobs for renderer noise, and opt-in refresh via `OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES=1`.
- Inspection snapshot/diff API for deterministic structure review across pages, masters, shapes, connectors, Shape Data, User cells, semantic tags, shape connection points, and connector waypoints.
- Stencil profile API for auditing whether diagrams are using generated masters, package-backed external masters, or plain geometry, with stable text output for regression review, aliases/keywords/tags, icon identity, placed/source dimension ranges, connection-point richness summaries, and persisted package-backed provenance after reload.
- Architecture, network, flowchart, block-diagram, swimlane, timeline, and sequence builder components now use first-party stencil catalogs for provenance, so cloud architecture, network segmentation, print audit trail, technical topology, governed process, release timeline, and incident sequence gallery output are baseline-reviewed as stencil-backed diagram content rather than anonymous geometry.
- Sequence diagrams now include first-party activation bars with dedicated builder API, stencil profile metadata, and gallery coverage.
- Stencil placement now stamps stencil id/name/category/catalog/source package/tags and package preview-image relationship metadata into shape and master metadata, so profiles and inspection snapshots can prove which catalog, package, and embedded icon media supplied a shape after save/load.
- Package stencil catalogs now extract native master connection points, persist them in catalog manifests, and scale them onto placed shapes so package-backed stencils expose usable connector attachment profiles after save/load.
- Generated stencil master instances now emit Visio-friendly page references by keeping `Master` and local style deltas while omitting generated `MasterShape` references unless a loaded shape explicitly preserved one.
- Obstacle-aware orthogonal routing APIs plus `PolishDiagram` options for rerouting connectors around unrelated top-level shapes, containers, background zones/trust boundaries, generated adornments, and existing connector paths before label cleanup, including deterministic page-level optimization passes and multi-waypoint dogleg candidates for crossing-aware routing sweeps.
- Group-aware obstacle routing can include nested group children while ignoring endpoint ancestors/descendants, so connectors into grouped content can avoid unrelated sibling member shapes.
- Zone-aware and connector-path-aware connector label cleanup that can prefer common endpoint zones, avoid unrelated background surfaces, move labels away from unrelated connector paths, and run whole-page optimization passes that revisit the most conflicted labels first during deterministic label placement.
- Lifeline-aware and fragment-label-aware normal sequence message label placement now prefers gaps between participant lifelines, activation bands, and fragment guard/adornment labels, with premium incident-sequence rendered and inspection baseline proof.
- Premium style presets for enterprise, technical, cloud, process, print-safe, and dark-safe diagrams, with validated technical/print smoke documents and rendered gallery baseline usage for the current market-facing set.

## Immediate P0

These are the items that still block "premium Visio library" positioning.

1. **Keep the public docs matched to the real API.**
   Website and product docs should use `VisioDocument`, diagram builders, and stencil catalogs, not older pseudo-APIs.

2. **Make examples visibly better.**
   The library now supports richer output, but the gallery needs stronger art direction:
   - fewer plain rectangles;
   - more external/native stencil usage;
   - better spacing, hierarchy, and caption treatment;
   - consistent typography;
   - restrained but polished themes;
   - examples that show actual business value, not only API coverage.

3. **Harden the showcase proof workflow.**
   Keep structural package validation, write a machine-readable showcase summary, and make optional desktop Visio SVG/PNG exports easy to review when Visio is available.

4. **Keep visual baselines meaningful.**
   The premium gallery now has generated SVG/PNG baselines plus PNG diff artifacts on failure. Keep them reviewable, refresh only deliberately, and pair rendered drift with structural inspection output where useful. The target gallery and proof levels are tracked in [officeimo.visio.premium-showcase.md](./officeimo.visio.premium-showcase.md).

5. **Keep the license story clean.**
   `OfficeIMO.Visio` should continue to match the repository MIT license and NuGet package metadata. Do not reintroduce local package license text that conflicts with `PackageLicenseExpression`.

## P1: Premium Diagram Experience

Goal: generated diagrams should be credible without manual post-editing.

- Continue tuning the richer theme catalog with diagram-specific margins, typography, connector weights, and label rules.
- Keep builder-level `Theme(...)` presets for enterprise, technical, cloud, process, print, and dark-safe diagrams covered by package validation, and keep expanding the rendered baseline gallery beyond the current technical and print-safe scenarios.
- Add diagram-specific visual defaults:
  - architecture zones and trust boundaries;
  - network subnets, racks, and device groupings;
  - flowchart continuation and branch retry patterns;
  - swimlane phase/lane readability;
  - sequence nested fragments and overlapping fragment layout;
  - dependency graph critical-path highlighting.
- Add automatic label collision cleanup for dense graph diagrams.
- Continue connector label cleanup beyond the current connector-path-aware, path-position-aware, whole-page optimization, lifeline-aware, and fragment-label-aware placement with denser graph and deeper sequence edge-case strategies.
- Extend orthogonal routing beyond deterministic page-level sweeps and dogleg candidates into deeper global dense-page crossing minimization.
- Add deterministic "polish passes" that can be applied after any builder.

## P2: Stencil Platform

Goal: native and external stencil packs feel first-class, not like shortcuts.

- Preserve more imported master/package content for unsupported external masters.
- Improve master metadata extraction: categories, keywords, preview/icon relationship metadata, dimensions, connection points, and aliases. Initial package preview metadata now records image relationship id, target, content type, extension, and embedded byte length when package masters expose image relationships, controlled by `VisioStencilPackageLoadOptions.ExtractPreviewImageMetadata`; native master connection points are learned by `ExtractConnectionPointMetadata` and persisted through catalog manifests; callers can also extract embedded preview/icon payloads with `VisioStencilPackageCatalog.ExtractPreviewImages(...)`, write them with `ExtractPreviewImagesToDirectory(...)`, or generate a reviewable HTML preview-payload gallery with deterministic SVG thumbnail artifacts for browser-renderable payloads through `CreatePreviewGallery(...)`.
- Add package-backed master reuse without forcing generated fallback where a real master exists.
- Add first-party generated stencil packs for:
  - servers and devices: initial Infrastructure catalog added;
  - cloud generic: initial Cloud catalog added;
  - security and identity: initial Security and Identity catalog added;
  - containers/Kubernetes generic: initial Containers and Kubernetes catalog added;
  - data/platform services: initial Data and Platform catalog added;
  - collaboration/business process symbols: initial Collaboration and Business Process catalog added.
- Add a real stencil gallery document builder for catalog review and debugging. Initial support now creates overview/category-paginated `.vsdx` review documents with visible Shape Data rows for stencil id, catalog/category, master, search metadata, source package, preview image, and connection-point counts.
- Add custom stencil export, ideally to reusable package-backed stencil form when feasible.
- Broaden typed stencil profiles from catalog/category/source-pack/family, alias/keyword/tag, placed/source dimensions, icon identity, preview-image content type/extension, extracted package preview/icon payloads, reviewable preview-payload gallery artifacts with deterministic SVG thumbnails for browser-renderable payloads, generated paginated gallery `.vsdx` documents with Shape Data metadata, learned native connection-point metadata, and placed connection-point summaries into richer package-family and native metadata.
- Document native Visio stencil discovery paths and external pack usage patterns.

## P3: Real Graphs And Data-Driven Diagrams

Goal: users can feed real inventories, dependencies, or workflows into OfficeIMO.Visio.

- Add graph import helpers from simple node/edge records. Initial `VisioGraphNodeRecord` / `VisioGraphEdgeRecord` support imports data-driven graph nodes and edges, including stencil catalog lookup, Shape Data, hyperlinks, roots, directed/undirected edges, and stable generated edge ids.
- Add stable ID and diff-friendly regeneration guidance. Initial graph-record imports derive missing connector ids from endpoint ids and connector kind.
- Add graph clustering/grouping APIs. Initial `Cluster(...)` / `Clusters(...)` support renders semantic graph clusters as background zones with Shape Data and hyperlinks, and `VisioGraphClusterRecord` can be imported with node/edge records for inventory-driven diagrams.
- Add dependency cycle presentation instead of only rejection where the diagram type allows cycles.
- Add legends based on used node/edge types. Initial generic graph `Legend(...)` support derives node-kind and connector-kind entries from the actual graph, reserves header layout space, and marks legend samples/text as generated diagram adornments.
- Add data-driven examples:
  - Azure/application dependency map;
  - Active Directory identity/authentication flow: initial reusable `VisioGallery` and official `--visio-showcase` example added with security/identity stencils, trust-boundary clusters, auth token/control/data flows, Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - CI/CD pipeline and build-agent topology: initial reusable `VisioGallery` example added with stencil-backed node records, edge records, cluster records, Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - Kubernetes/service-mesh topology: initial reusable `VisioGallery` and official `--visio-showcase` example added with Kubernetes/data/cloud stencils, service-mesh clusters, mTLS/control/data flows, Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
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
- Use inspection diffs next to visual baseline failures so review output explains both structural and rendered changes.
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
