# OfficeIMO.Visio Roadmap

Date: 2026-06-05
Current branch/worktree: `codex/visio-market-review-20260605` at `C:\Support\GitHub\OfficeIMO-visio-market-review-20260605`
Current baseline: `origin/master` at `b31d0f80` (`Merge pull request #1884 from EvotecIT/codex/pdf-premium-next-20260604`)
Last major Visio PR merged: https://github.com/EvotecIT/OfficeIMO/pull/1865

## Where We Are

OfficeIMO.Visio is no longer just a basic VSDX writer. The current branch has a usable, dependency-light Visio authoring stack:

- VSDX package creation, loading, editing, saving, validation, and preservation of unsupported content.
- Pages, shapes, connectors, connection points, groups, layers, hyperlinks, User cells, typed Shape Data, protection, page settings, backgrounds, metadata, themes, style sheets, and master-backed page instances.
- Fluent page and document authoring for lower-level diagrams.
- Fluent loaded-page editing with typed page targeting, selection queries, replace-master helpers, and reusable stencil migration maps.
- High-level builders for flowcharts, block diagrams, dependency diagrams, architecture diagrams, networks, network topology, swimlanes, org charts, timelines, sequences, and generic graphs.
- Reusable style themes, premium enterprise/technical/cloud/process/print/dark-safe presets, and local node/edge style overrides.
- Connector routing, obstacle-aware routing around unrelated shapes, optional zone/container/group/adornment-aware routing, connector-crossing-aware route scoring, deterministic page-level routing optimization passes, multi-waypoint dogleg candidates for dense crossings, label placement, whole-page connector label optimization passes, zone-aware and connector-path-aware label cleanup, page fitting, deterministic text measurement through `OfficeIMO.Drawing`, and visual quality analysis.
- Header-style region/zone captions for architecture, block, network, topology, and graph builders, including layout clearance and quality-analyzer handling for generated caption adornments.
- First-party generated stencil catalogs and package-backed stencil catalogs.
- External `.vssx`, `.vstx`, and `.vsdx` catalog loading, including multi-package external stencil repositories.
- Native/installed Visio stencil discovery when Visio is available, without making Office automation part of the core library.
- Native headless SVG and managed PNG export for OfficeIMO-authored pages, plus optional desktop Visio validation/export helpers for local proof, separate from the dependency-free core.
- Generated showcase examples and preview output for visual inspection.
- Premium showcase diagrams with validated VSDX packages plus local PNG/SVG preview proof.
- A reusable eight-diagram `VisioPremiumGallery` plus approved PNG/SVG baseline fixtures and PNG diff artifacts for both representative native renderer proof and Visio desktop preview regression.
- Deterministic inspection snapshots and structural diffs through `CreateInspectionSnapshot()` / `VisioInspectionDiff`, including connector label placement coordinates for label-layout regression proof.
- Deterministic stencil usage profiles through `CreateStencilProfile()`, including generated-master, package-backed, basic-geometry, stencil-backed, Shape Data key, semantic-kind, catalog, category, keyword, alias, tag, icon identity, package preview-image content type/extension, learned native connection points, source default dimensions/units, source-package, placed/source-dimension, and connection-point summaries that survive save/load for package-backed masters and generated stencil placements.
- Deterministic visual-quality proof text through `VisioDiagramQualityReport.ToText()`, including `quality.*` issue counts and per-issue fields that can travel with generated showcase artifacts, plus parsed clean/issue/error/warning/information rollups in `VisioShowcaseSummary`.

The external stencil and graph slice from PR #1865 is merged. The next checkpoint is no longer proving that the Visio core can generate useful diagrams; it is making the public package story clean, the showcase reproducible, and the generated output visually strong enough to support premium positioning.

## 2026-06-04 Checkpoint

Current focused proof on this worktree:

```powershell
dotnet test .\OfficeIMO.Visio.Tests\OfficeIMO.Visio.Tests.csproj -c Release --framework net8.0 --filter "FullyQualifiedName~Visio"
```

Result: `738/738` Visio-filtered tests passed.

Issues fixed in this checkpoint:

- Data-graphic bars now reject non-finite or reversed numeric ranges before generating geometry, preventing invalid or silently misleading premium data adornments.
- Fluent authoring can now target loaded pages without adding duplicates through `FirstPage(...)`, `ExistingPage(...)`, `Page(VisioPage, ...)`, and idempotent `PageOrAdd(...)`.
- Fluent loaded-page editing can now bulk-edit typed shape and connector selections through `Shapes(...)`, `ShapesWithData(...)`, `ShapesWithShapeData(...)`, `ShapesContainingText(...)`, `ShapesInLayer(...)`, and `Connectors(...)`.
- Advanced fluent loaded-page editing now exposes geometry, graph, metadata, and connector-neighborhood selectors including `ShapesIntersecting(...)`, `ShapesContainedIn(...)`, `ConnectedComponent(...)`, `PathBetween(...)`, `ShapesWithUserCell(...)`, `ShapesWithHyperlink(...)`, `ShapesWithProtection(...)`, `OutgoingConnectors(...)`, `IncomingConnectors(...)`, `ConnectedConnectors(...)`, `ConnectorsInLayer(...)`, `ConnectorsWithHyperlink(...)`, and `ConnectorsWithProtection(...)`.
- Fluent loaded-page editing now exposes replace-master and stencil-standardization workflows through `ReplaceMaster(...)`, `ReplaceMasters(...)`, and `ReplaceMastersByMaster(...)`, backed by the existing page/selection master editing engine.
- Typed stencil migration maps now standardize loaded documents, pages, and page-backed selections by current stencil id, master name, shape `NameU`, or typed predicate in one first-match-wins pass, with fluent `ApplyStencilMigration(...)` support and saved/reloaded validation coverage.
- Migration maps can now resolve replacement shapes from first-party or package-backed catalogs by query, and `VisioStencilMigrationPresets.BasicFlowchart(...)` upgrades unstenciled/basic flowchart-like diagrams to semantic catalog stencils without touching shapes that already carry OfficeIMO stencil metadata.
- `PlanStencilMigration(...)` now provides a non-mutating dry-run for documents, pages, and page-backed selections, with `VisioStencilMigrationPlan.ToText()` producing stable review/CI/operator reports before callers apply a migration.
- Reviewed migration plans can now be replayed through `ApplyStencilMigration(plan, map)` on documents, pages, page-backed selections, and fluent document/page chains. The apply path validates the planned page identity, shape id, original text, original master/stencil metadata, match rule, and replacement stencil before it mutates, so approved reports fail closed when the diagram or map drifts.
- Migration plans can now be persisted as dependency-light text artifacts through `SaveText(...)`, `LoadText(...)`, and `FromText(...)`, preserving explicit null/empty text state so approval can happen between processes without weakening drift validation.
- `VisioStencilMigrationPresets.NetworkInfrastructure(...)` and `ArchitectureInfrastructure(...)` now upgrade common unstenciled/labeled network, infrastructure, and architecture shapes to semantic catalog stencils while preserving already-stenciled content.
- `VisioStencilMigrationPresets.OrgChart(...)`, `Timeline(...)`, and `Sequence(...)` now upgrade common unstenciled/labeled organization chart, roadmap/timeline, and sequence diagram shapes to semantic catalog stencils with conservative text/name matching and already-stenciled skip behavior.
- `VisioStencilMigrationPresets.SwimlaneProcessMap(...)`, `CloudInfrastructure(...)`, and `SecurityIdentity(...)` now cover common lane/phase/activity, cloud boundary/region/service/function/queue, and identity/security policy/firewall/audit/alert cleanup with the same conservative skip behavior.
- `VisioStencilMigrationPresets.ContainersKubernetes(...)`, `DataPlatform(...)`, and `CollaborationBusiness(...)` now cover common Kubernetes/container, data platform, and collaboration/business-process cleanup, including a fixed container-stencil query that avoids incorrectly choosing pod stencils for container-image shapes.
- Same-page duplication now has typed `VisioShapeDuplicationOptions` for offsets, internal connector copying, semantic shape/connector ID suffixes, and advanced ID factories; fluent loaded-page editing exposes `DuplicateShape(...)` and `DuplicateShapes(...)` with friendly `-copy` IDs so copied shapes remain chain-addressable.
- Native containers can now be maintained in load-edit-save workflows through typed membership APIs: `AddToContainer(...)`, `RemoveFromContainer(...)`, `RefitContainer(...)`, `GetContainerMembers(...)`, selection `WrapInContainer(...)`, and fluent ID-based container editing update the modeled relationship graph and clear stale loaded formulas before saving fresh Visio `DEPENDSON(...)` relationships.
- Native container metadata and styles can now be inspected and updated after load through `VisioContainerInfo`, `GetContainerInfo(...)`, `GetContainerOptions(...)`, `ApplyContainerOptions(...)`, `ConfigureContainer(...)`, and fluent `ContainerInfo(...)` / `ConfigureContainer(...)` / `StyleContainer(...)` helpers. The update path writes Visio-native container User cells for margin, resize, lock, no-highlight/no-ribbon, style ids, and heading style, preserves custom heading height through an OfficeIMO User cell, and fixes metric-page refit so stored inch-based margins are converted back to page units before refit.
- Native VSDX comments now round-trip through a typed `VisioComment` model, page/shape-targeted `AddComment(...)` / `AddCommentToShape(...)` helpers, typed review verbs (`UpdateText(...)`, `Resolve(...)`, `Reopen(...)`, `UpdateCommentText(...)`, `ResolveComment(...)`, `ReopenComment(...)`, `RemoveComment(...)`), fluent `Comment(...)` / `CommentShape(...)` / `UpdateComment(...)` / `ResolveComment(...)` / `ReopenComment(...)` loaded-page verbs, and real `/visio/comments.xml` package author/comment entries with persisted PageID/ShapeID, `Done`, and `EditDate` references.
- Swimlane diagrams now persist typed lane, phase, and activity placement metadata and can be maintained after load through `GetSwimlaneLanes(...)`, `GetSwimlanePhases(...)`, `GetSwimlaneActivities(...)`, `MoveSwimlaneActivity(...)`, `RelayoutSwimlaneActivities(...)`, and fluent `MoveSwimlaneActivity(...)` / `RelayoutSwimlanes(...)` verbs that infer older generated layouts, stack destination cells, and reroute affected connectors.
- Loaded-page fluent editing now exposes deterministic relayout workflows through `RelayoutShapesAsGrid(...)`, `RelayoutShapesAsHorizontalStack(...)`, `RelayoutShapesAsVerticalStack(...)`, `RelayoutConnectedComponentAsGrid(...)`, and `RelayoutContainerMembers(...)`, all backed by typed `VisioShapeSelection` layout helpers; selection relayout now rejects non-finite spacing before it can poison generated coordinates.
- `OfficeIMO.Visio` README no longer describes the package as early/minimal after documenting premium diagrams, stencil catalogs, native SVG/PNG export, and visual proof.
- The assessment now reflects the current branch and the resolved MIT license state.

Next best slice:

- Continue expanding editing workflows from migration, duplication, container membership/style maintenance, native comment review, first-pass swimlane maintenance, and fluent relayout helpers into richer comment threading/author workflows, deeper swimlane metadata/auto-assignment, richer nested/container behavior, and broader diagram-level relayout/polish workflows.
- Keep no-dependency SVG/PNG export as a first-class proof path and add more renderer fidelity cases only when the generated gallery exposes real visual drift.
- Move public docs and website examples to semantic builders, stencil catalogs, `VisioPremiumGallery`, and native preview APIs.
- Continue replacing anonymous geometry in market-facing examples with first-party or package-backed stencils where the domain benefits from recognizable symbols.
- Continue from the new artifact-friendly showcase output by wiring `showcase-summary.json`, `showcase-gallery.html`, native SVG/PNG preview galleries, structural inspection/profile/visual-quality proof artifacts, top-level proof/evidence totals, diagram-level package/preview/proof/evidence records, complete review proof, shape/connector and Shape Data proof rollups, stencil-backed/basic-geometry mix, connection-point coverage, stencil provenance/coverage, and the manual `[self-hosted, Windows, Visio]` desktop proof lane into CI/PR artifacts.

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
- Sequence diagrams now include first-party activation bars, guarded/partitioned fragments, nested/overlapping fragment layout metadata, dedicated builder API, stencil profile metadata, and gallery coverage.
- Stencil placement now stamps stencil id/name/category/catalog/source package/tags and package preview-image relationship metadata into shape and master metadata, so profiles and inspection snapshots can prove which catalog, package, and embedded icon media supplied a shape after save/load.
- Package stencil catalogs now extract native master connection points, persist them in catalog manifests, and scale them onto placed shapes so package-backed stencils expose usable connector attachment profiles after save/load.
- Generated stencil master instances now emit Visio-friendly page references by keeping `Master` and local style deltas while omitting generated `MasterShape` references unless a loaded shape explicitly preserved one.
- Shape Data data graphics now render as local badge/bar geometry even in master-backed documents, and the premium executive dependency gallery baseline promotes visible `Status` and `SLO` adornments with approved PNG/SVG plus inspection/profile proof.
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
   Keep structural package validation, use the generated `showcase-summary.json` machine-readable artifact, top-level proof/evidence totals, and diagram-review lists, expose `showcase-gallery.html` headline proof/evidence metrics, evidence coverage, and diagram review cards for quick visual review, inspection/profile/visual-quality proof links, complete review proof, shape/connector and Shape Data proof rollups, stencil-backed/basic-geometry mix, connection-point coverage, and use the manual `[self-hosted, Windows, Visio]` desktop lane for Microsoft Visio SVG/PNG exports when Visio is available.

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
  - sequence nested fragments and overlapping fragment layout: initial `NestedFragment(...)` and nested `VisioSequenceFragmentRecord.ParentFragmentId` support now inset child/overlapping frames, stamps parent/depth/overlap-lane metadata, and is promoted through the incident/runbook gallery;
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
- Add reusable migration-map presets that can standardize common generated/basic shape families into first-party or package-backed catalog shapes. Initial catalog-query rules, apply-from-plan validation, persisted text plan artifacts, and basic-flowchart, network/infrastructure, architecture, org-chart, timeline, sequence, swimlane/process-map, cloud infrastructure, security/identity, Kubernetes/container, data/platform, and collaboration/business presets are in place; continue only with additional domain presets where labels are reliable enough.
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
- Add sequence import helpers from simple participant/message/activation/fragment/note records. Initial `VisioSequence*Record` support imports data-driven runbook sequences with participants, self-messages, activations, guarded/nested fragments, partitions, parent-fragment metadata, and semantic notes.
- Add network import helpers from simple zone/node/link/callout records. Initial `VisioNetwork*Record` support imports data-driven network segmentation diagrams with Shape Data, hyperlinks, stable link ids, zones, stencil-backed nodes, and semantic callouts.
- Add stable ID and diff-friendly regeneration guidance. Initial graph-record imports derive missing connector ids from endpoint ids and connector kind.
- Add graph clustering/grouping APIs. Initial `Cluster(...)` / `Clusters(...)` support renders semantic graph clusters as background zones with Shape Data and hyperlinks, and `VisioGraphClusterRecord` can be imported with node/edge records for inventory-driven diagrams.
- Add dependency cycle presentation instead of only rejection where the diagram type allows cycles.
- Add legends based on used node/edge types. Initial generic graph `Legend(...)` support derives node-kind and connector-kind entries from the actual graph, reserves header layout space, and marks legend samples/text as generated diagram adornments.
- Add data-driven examples:
  - Azure/application dependency map: initial reusable `VisioGallery` and official `--visio-showcase` example added with cloud/security/data/collaboration stencils, edge-security and application-runtime clusters, inspectable Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - Active Directory identity/authentication flow: initial reusable `VisioGallery` and official `--visio-showcase` example added with security/identity stencils, trust-boundary clusters, auth token/control/data flows, Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - privileged access review graph: initial reusable `VisioGallery` and official `--visio-showcase` example added with security/identity, cloud, infrastructure, and collaboration stencils, access-control and evidence clusters, policy/session/audit Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - data-platform lineage graph: initial reusable `VisioGallery` and official `--visio-showcase` example added with data/platform, cloud, and collaboration stencils, ingestion/quality and serving/governance clusters, CDC/batch/lineage/query flows, Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - hybrid network operations graph: initial reusable `VisioGallery` and official `--visio-showcase` example added with network, infrastructure, cloud, data/platform, and collaboration stencils, edge/WAN, datacenter rack/service, and operations-review clusters, backup/rack/server/storage metadata, stable connector ids, validation, quality analysis, and automatic graph legend coverage;
  - process governance review graph: initial reusable `VisioGallery` and official `--visio-showcase` example added with collaboration/business, flowchart, and security stencils, intake/assessment, governance/execution, and evidence/audit clusters, stable connector ids, Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - CI/CD pipeline and build-agent topology: initial reusable `VisioGallery` example added with stencil-backed node records, edge records, cluster records, Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - Kubernetes/service-mesh topology: initial reusable `VisioGallery` and official `--visio-showcase` example added with Kubernetes/data/cloud stencils, service-mesh clusters, mTLS/control/data flows, Shape Data, hyperlinks, validation, quality analysis, and automatic graph legend coverage;
  - incident/runbook sequence: initial reusable `VisioGallery` and official `--visio-showcase` example added with sequence record imports, actor/control/entity/database participants, activations, guarded/nested recovery fragments, runbook notes, validation, quality analysis, and sequence-stencil profile coverage;
  - network segmentation diagram: initial reusable `VisioGallery` and official `--visio-showcase` example added with network record imports, segmented zones, stencil-backed network nodes, Shape Data, hyperlinks, stable link ids, callouts, validation, quality analysis, and network-stencil profile coverage.

## P4: Editing Existing Diagrams

Goal: OfficeIMO can safely update diagrams created elsewhere.

- Expand replace-master operations for loaded external masters.
- Add richer selection queries: geometry intersection, contained-in-zone, connected component, path search, and data predicates. Initial query support now includes `ShapesIntersecting(...)`, `ShapesContainedIn(...)`, Shape Data predicate queries, `ConnectedComponent(...)`, and `PathBetween(...)` with matching editable selection helpers. Fluent loaded-document editing now exposes page targeting, bulk typed shape/connector selection helpers, id-based geometry/path/component selectors, connector-neighborhood selectors, replace-master helpers, typed stencil migration maps, catalog-query migration rules, basic-flowchart/network-infrastructure/architecture/org-chart/timeline/sequence/swimlane/cloud/security/Kubernetes/data/collaboration migration presets, non-mutating migration plans/reports, persisted migration-plan text artifacts, and validated apply-from-plan workflows for common create-or-edit and load-edit-save workflows.
- Add comment APIs and richer annotation/callout editing. Initial native comment support now saves, loads, fluently edits, updates, resolves, reopens, filters, and removes page-level and shape-targeted Visio comments as real `/visio/comments.xml` package content with typed authors, timestamps, done state, edit dates, and persisted shape references.
- Add containers and swimlanes as deeper typed concepts. Initial native container create/load/query support now includes typed add/remove/refit membership editing and fluent ID-based container maintenance for loaded pages.
- Add data graphics and richer Shape Data schema helpers. Initial schema support now provides reusable `VisioShapeDataSchema` / `VisioShapeDataField` definitions with defaults, labels, prompts, types, list formats, sort keys, required-value validation, allowed-value validation, and bulk application to shapes, connectors, and selections. Initial data graphic support now turns Shape Data values into generated badge/bar adornments linked back to the target shape, field, value, and role, saves those generated visuals as renderer-friendly local geometry in master-backed documents, and is promoted through the premium executive dependency baseline.
- Add safe relayout of selected subsets while preserving unsupported content.
- Add round-trip tests against more Visio-authored assets.

## P5: Verification And Export

Goal: prove quality continuously.

- Initial first-party SVG preview/export is available through `ToSvg()` and `SaveAsSvg(...)` without requiring Microsoft Visio. SVG text now uses bounded wrapping/scaling for shape text blocks, including long-word breaks, alpha-preserving styled text colors, styled underline/italic attributes and authored text rotation, rotated styled text-block backgrounds with Visio transparency, readable connector-label backgrounds, color/opacity-matching inline connector arrowheads, non-mutating render-time connector-label overlap avoidance including endpoint-shape, dense-label clearance, and connector-line crossing avoidance, metadata-driven first-party stencil pictograms with authored shape rotation for straight and curved glyphs for parity with the native PNG preview path, semantic database/storage cylinder bodies while plain flowchart `Data` remains a parallelogram, semantic flowchart start/end terminator capsules, semantic document stencil wavy bottoms, built-in chevron polygons, delay D-shapes, and manual input slanted quadrilaterals, simple preserved Visio `MoveTo`/`LineTo`/`PolylineTo` and relative geometry-row outlines with intra-section subpath breaks and unclosed NoFill open paths, deleted Geometry/SplineKnot rows, simple preserved `Width`/`Height`/`LocPinX`/`LocPinY`/`PinX`/`PinY`/`Angle`/`MIN`/`MAX`/`ABS`/`SQRT`/`PI`/`SIN`/`COS`/`TAN`/`ATAN`/`ATAN2`/`RAD`/`DEG`/`INT`/`POW`/`^`/`ROUND`/`AND`/`OR`/`NOT`/`GUARD`/`IF` formulas including `POLYLINE` arguments and percentage plus angle/length unit-suffixed numeric literals, preserved `NoFill`/`NoLine`/`NoShow` geometry flags, scaled master/master-shape preserved outlines, preserved `Ellipse` and open clipped `InfiniteLine` rows, flattened preserved `ArcTo`/`EllipticalArcTo`, `RelEllipticalArcTo`, `CubBezTo`/`QuadBezTo`, `RelCubBezTo`/`RelQuadBezTo`, `SplineStart`/`SplineKnot`, and `NURBSTo` formula outlines with Visio compact knot-vector expansion, and rotated browser-renderable package-backed preview/icon payload projection including content-sniffed generic media relationships when package masters expose embedded PNG/JPG/GIF/SVG media. Initial managed, dependency-free native PNG export is available through `ToPng()` and `SaveAsPng(...)` for OfficeIMO-authored geometry; it must stay free of operating-system graphics/font APIs. PNG text now uses managed TrueType/OpenType outline contours from `OfficeIMO.Drawing` with explicit `FontFilePath`/`FontCollectionIndex`/`FontFaceName` controls or managed default font-file discovery, with the small stroke font only as fallback, preserves bounded long-word wrapping, alpha-blended styled text colors, styled underlines, oblique styled italics, authored text rotation, and rotated styled text-block backgrounds, projects the same first-party stencil metadata as dependency-free vector/raster pictograms with authored shape rotation for straight and curved glyphs, renders rotated ellipse shapes, semantic database/storage cylinder bodies, semantic flowchart start/end terminator capsules, semantic document stencil wavy bottoms, and built-in chevron polygons, delay D-shapes, and manual input slanted quadrilaterals plus the same simple preserved Visio `MoveTo`/`LineTo`/`PolylineTo`, relative geometry-row, intra-section subpath breaks, and unclosed NoFill open paths, deleted Geometry/SplineKnot rows, simple `Width`/`Height`/`LocPinX`/`LocPinY`/`PinX`/`PinY`/`Angle`/`MIN`/`MAX`/`ABS`/`SQRT`/`PI`/`SIN`/`COS`/`TAN`/`ATAN`/`ATAN2`/`RAD`/`DEG`/`INT`/`POW`/`^`/`ROUND`/`AND`/`OR`/`NOT`/`GUARD`/`IF` formula including `POLYLINE` arguments and percentage plus angle/length unit-suffixed numeric literals, preserved `NoFill`/`NoLine`/`NoShow` geometry flags, scaled master/master-shape, preserved `Ellipse` and open clipped `InfiniteLine` rows, and flattened `ArcTo`/`EllipticalArcTo`, `RelEllipticalArcTo`, `CubBezTo`/`QuadBezTo`, `RelCubBezTo`/`RelQuadBezTo`, `SplineStart`/`SplineKnot`, plus `NURBSTo` formula outlines with Visio compact knot-vector expansion, preserves dashed shape and connector strokes, projects embedded package PNG preview/icon payloads including content-sniffed generic media relationships, truecolor, indexed-palette, grayscale, and grayscale-alpha PNGs, including packed 1/2/4-bit indexed/grayscale icons, 16-bit channel payloads downsampled into the native preview buffer, palette and truecolor tRNS transparency, aspect-preserving placement, and shape rotation, uses the same render-time connector-label overlap avoidance including endpoint-shape, dense-label clearance, and connector-line crossing avoidance, and writes managed DEFLATE-compressed PNG output rather than stored blocks. The no-Visio native SVG/PNG baseline lane now covers all eight premium gallery diagrams. The next export work is broader package-backed artwork beyond embedded PNG/browser-renderable previews, broader NURBS periodic/open-edge cases, broader shape geometry, and deeper dense-graph connector-label collision handling.
- Use `OfficeIMO.Pdf` and `OfficeIMO.Drawing` for optional previews where they fit, without making them the VSDX source of truth.
- Expand desktop Visio validation to collect repair dialogs, export failures, and visual artifacts.
- CI now has a focused Visio showcase artifact workflow for generated packages, top-level proof/evidence totals and diagram-level package/preview/proof/evidence records in versioned `showcase-summary.json`, `showcase-gallery.html`, structural inspection/profile/visual-quality proof artifacts, and native SVG/PNG outputs from `--visio-showcase --visio-native-preview`; the reusable `VisioShowcaseSummary` validation API checks generated artifacts, recomputed proof totals, parsed visual-quality rollups, and recomputed evidence totals before summary publication, the shared workflow validator scripts independently check `schemaVersion`, summary counts, `artifactCount`, `proofTotals`, `evidenceTotals`, package-to-preview/proof grouping, native/desktop preview-format completeness, complete review proof, visual-quality proof schema and rollups, stencil proof summaries, stencil catalog coverage, artifact existence, file sizes, recomputed SHA-256 hashes, gallery review-index deep links, gallery links/full-hash evidence, and Markdown artifact fingerprints before upload, and the workflow also exposes a manual `run_desktop_preview` dispatch lane for a `[self-hosted, Windows, Visio]` runner where Microsoft Visio desktop is installed.
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
