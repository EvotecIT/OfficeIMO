# OfficeIMO.Visio Premium Showcase Plan

Date: 2026-06-01

## Goal

The showcase should prove that OfficeIMO.Visio can generate diagrams that are structurally valid, easy to edit in Visio, and visually credible without hand placement.

## Proof Levels

1. **Package proof**
   - Run `--visio-showcase`.
   - Validate every generated `.vsdx` with `VisioValidator`.
   - Validate generated showcase artifacts with `VisioShowcaseSummary.EnsureArtifactsValid(...)`.
   - Write deterministic inspection, stencil-profile, and visual-quality text artifacts under `Documents/Visio Showcase/Structural Proof`.
   - Write `showcase-summary.md`, `showcase-summary.json`, and `showcase-gallery.html` with generated package, preview, review proof, artifact size/hash, `schemaVersion`, `artifactCount`, top-level `proofTotals`, top-level `evidenceTotals`, diagram-level package/preview/proof/evidence metadata, structural shape/connector and Shape Data rollups, stencil-backed/basic-geometry mix, connection-point coverage, visual-quality `quality.*` counts, stencil catalog proof summaries parsed from `.stencil-profile.txt` artifacts, and stencil catalog coverage grouped by diagram.
   - Use the HTML header metrics, evidence coverage table, proof totals, diagram review cards, and stencil coverage table to inspect the overall generated native-preview/complete-review-proof completeness, shape/connector/stencil/connection-point coverage, and each generated `.vsdx` beside its available native or desktop previews, review proof links, SHA-256 fingerprints, shape/connector metrics, Shape Data counts, stencil backing metrics, connection-point counts, visual-quality issue counts, and stencil catalog chips.

2. **Native preview proof**
   - Run `--visio-showcase --visio-native-preview`.
   - Export first-page SVG and PNG previews through the dependency-free renderers.
   - Write the native preview gallery under `Documents/Visio Showcase/Native Preview`.
   - Include native SVG/PNG paths, review proof paths, formats, sizes, SHA-256 hashes, artifact kind, package-to-preview/proof grouping, top-level proof totals, top-level evidence totals, structural shape/connector and Shape Data metrics, stencil-backed/basic-geometry mix, connection-point coverage, visual-quality evidence, diagram stencil provenance, and catalog coverage in `showcase-summary.json`, with quick visual review through `showcase-gallery.html`.
   - Treat this as a fast CI/review lane for OfficeIMO-authored geometry, not as a substitute for desktop Visio visual parity.

3. **Desktop proof**
   - Run `--visio-showcase --visio-preview` on a machine with Microsoft Visio.
   - Optional desktop round-trip and PNG/SVG preview checks live in the test/manual baseline suite, not the shipped `OfficeIMO.Visio` package.
   - Add preview files and gallery output as review artifacts.
   - Use the `Visio Showcase Artifacts` workflow through `workflow_dispatch` with `run_desktop_preview=true` on a `[self-hosted, Windows, Visio]` runner where Microsoft Visio desktop automation is registered.

4. **Visual baseline proof**
   - Store approved PNG/SVG previews for the core premium gallery under `OfficeIMO.Visio.Tests/Visio/VisualBaselines`.
   - Run a no-Visio native baseline lane for all eight premium diagrams; it compares first-party `ToSvg()` and `SaveAsPng(...)` output so renderer drift is gated even on machines without Microsoft Visio.
   - Store approved `.inspection.txt` and `.stencil-profile.txt` snapshots for the same gallery so rendered drift can be reviewed beside structural and stencil-usage drift.
   - Showcase artifact runs additionally publish `.visual-quality.txt` proofs with deterministic `quality.*` counts; promote those into approved premium baselines once the baseline suite starts tracking quality drift directly.
   - Compare regenerated previews, inspection snapshots, and stencil profiles with `VisioPremiumVisualBaselineTests`.
   - Refresh approved artifacts deliberately with `OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES=1`.
   - Refresh only inspection/profile baselines with `OFFICEIMO_UPDATE_VISIO_PREMIUM_STRUCTURAL_BASELINES=1`.
   - Run the desktop-rendered baseline lane explicitly with `OFFICEIMO_RUN_VISIO_PREMIUM_DESKTOP_BASELINES=1`, or require it with `OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES=1`; use `OFFICEIMO_RUN_VISIO_DESKTOP_VALIDATION=1` for the smaller desktop validator smoke tests. Otherwise the native baseline lane still runs and Microsoft Visio COM open/export checks are kept out of default local/CI test runs.
   - On PNG drift, write expected, actual, and `.diff.png` artifacts with changed-pixel count and max channel delta, plus inspection/profile expected, actual, and `.diff.txt` artifacts for the same diagram.
   - On structural drift, write inspection/profile expected, actual, and `.diff.txt` artifacts.
   - Use `OFFICEIMO_VISIO_PREMIUM_BASELINE_PIXEL_TOLERANCE` and `OFFICEIMO_VISIO_PREMIUM_BASELINE_ALLOWED_DIFF_PIXELS` only for deliberate renderer-noise tolerance.
   - Treat structural validity and visual drift as separate gates.

## Premium Gallery Targets

- **Cloud architecture:** first-party architecture-stencil zones, trust boundaries, gateway, eventing, data store, runbook links, Shape Data.
- **Network segmentation:** first-party network-stencil internet, firewall, core switch, server zone, client zone, monitoring, labeled links.
- **Executive dependency graph:** layered first-party architecture stencil graph with owners, critical path, external dependencies, and grouped runtime zone.
- **Privileged access review:** first-party security, cloud, infrastructure, and collaboration stencils for just-in-time access approval, policy evaluation, session brokerage, secret checkout, privileged target access, and evidence review.
- **Technical topology:** first-party block-diagram stencil runtime mesh with edge/runtime/state zones, policy path, eventing, cache, and secrets flow.
- **Print audit trail:** first-party flowchart-stencil access-review flow with monochrome-safe typography and compliance-friendly spacing.
- **Incident sequence:** first-party sequence-stencil actor/system/service participants, activation bars, guarded/partitioned combined fragment frame, collision-aware note placement, self-message placement, notes API coverage, error path, and remediation path; still needs deeper nested and multi-fragment layout polish.
- **Release timeline:** first-party timeline-stencil milestones, risks, spans, decisions, callouts, and status metadata.
- **Swimlane process:** first-party swimlane-stencil roles, phases, handoffs, exception path, compliance callouts, and readable lane headers.

## Design Bar

- Use semantic builders first; drop to low-level shapes only for deliberate visual accents.
- Prefer real first-party or package-backed stencil nodes where the domain benefits from recognisable symbols.
- Use stable IDs, hyperlinks, and Shape Data so diagrams are searchable and regeneration-friendly.
- Use the premium enterprise, technical, cloud, process, print-safe, and dark-safe preset set as the default theme catalog; promote only visually reviewed scenarios into the rendered gallery baseline.
- Make every diagram explain a real business or operational scenario, not just API coverage.

## Next Implementation Slice

1. Extend obstacle-aware routing beyond deterministic page-level optimization sweeps and multi-waypoint dogleg candidates into deeper global dense-page minimization; connector labels now avoid unrelated connector paths, can slide along their connector paths before offsetting, and can run whole-page optimization passes before normal sequence message lifeline-aware and fragment-label-aware placement, so the next label work is denser graph and deeper sequence edge-case cleanup.
2. Keep replacing anonymous basic geometry with stencil-backed premium symbols where the domain supports it; all eight current premium baselines now carry first-party stencil provenance for their domain shapes, the native SVG/PNG preview path now projects built-in stencil metadata into dependency-free rotated pictograms, and the reusable catalog set now includes first-party infrastructure, cloud, security/identity, Kubernetes, data/platform, and collaboration/business packs ready for deeper gallery promotion. Sequence activations, guarded/partitioned combined fragments, and collision-aware notes are promoted while deeper nested sequence fragments still need stronger layout.
3. Keep tightening the baseline-reviewed technical and print-safe scenarios before using those presets in website screenshots or product screenshots.
4. Extend current premium inspection/profile baselines and showcase quality proofs beyond structural shape/connector rollups, visual-quality `quality.*` counts, top-level showcase proof and evidence totals, Shape Data and connector Shape Data key counts, semantic-kind counts, stencil-backed/basic-geometry mix, connection-point coverage, typed family, alias/keyword/tag, placed/source-dimension, icon identity, preview-image relationship metadata, embedded preview/icon payload extraction, reviewable HTML preview-payload galleries with deterministic SVG thumbnails for browser-renderable payloads, generated paginated `.vsdx` stencil-gallery review documents with Shape Data metadata, learned native connection points, and placed connection-point rollups into richer package-family details and deeper native-pack metadata as external metadata extraction improves.
5. Keep all eight native SVG/PNG premium baselines useful by fixing renderer artifacts before approving drift; alpha-preserving SVG/PNG styled text colors, native PNG long-word wrapping, native PNG styled text underlines/italics/TextAngle rotation, rotated styled text-block backgrounds, rotated package-backed preview artwork, rotated first-party stencil pictograms including curved glyphs, semantic database/storage cylinder bodies while plain flowchart `Data` remains a parallelogram, semantic flowchart start/end terminator capsules, semantic document stencil wavy bottoms, built-in chevron polygons, delay D-shapes, and manual input slanted quadrilaterals, native PNG rotated ellipse shapes, color/opacity-matching inline SVG connector arrowheads, package-backed PNG/browser-renderable preview payload projection including content-sniffed generic media relationships, native PNG tRNS transparency handling, simple preserved Visio `MoveTo`/`LineTo`/`PolylineTo` plus relative geometry-row outlines with intra-section subpath breaks and unclosed NoFill open paths, deleted Geometry/SplineKnot rows, simple preserved `Width`/`Height`/`LocPinX`/`LocPinY`/`PinX`/`PinY`/`Angle`/`MIN`/`MAX`/`ABS`/`SQRT`/`PI`/`SIN`/`COS`/`TAN`/`ATAN`/`ATAN2`/`RAD`/`DEG`/`INT`/`POW`/`^`/`ROUND`/`AND`/`OR`/`NOT`/`GUARD`/`IF` formulas including `POLYLINE` arguments and percentage plus angle/length unit-suffixed numeric literals, preserved `NoFill`/`NoLine`/`NoShow` geometry flags, scaled master/master-shape preserved outlines, preserved `Ellipse` and open clipped `InfiniteLine` rows, flattened preserved `ArcTo`/`EllipticalArcTo`, `RelEllipticalArcTo`, `CubBezTo`/`QuadBezTo`, `RelCubBezTo`/`RelQuadBezTo`, `SplineStart`/`SplineKnot`, `NURBSTo` formula outlines with Visio compact knot-vector expansion, native PNG dashed shape/connector strokes, and render-time connector-label avoidance for endpoint shapes, dense-label clearance, and connector-line crossings are now supported renderer lanes, and the next fidelity work is broader package-backed artwork beyond those previews, broader NURBS periodic/open-edge cases, broader shape coverage, and deeper dense-graph label routing.
6. Broaden the gallery from eight baseline-approved diagrams toward a larger market-facing set; the newly promoted reusable scenarios include a privileged-access review graph with policy, PAM, vault, target, audit, and SIEM evidence flow, a data-platform lineage graph with ingestion, quality, lake, warehouse, catalog, query API, and analytics flow, a hybrid network operations graph with edge/WAN, datacenter racks, compute hosts, storage, backup flow, operations monitoring, and NOC review flow, and a process governance graph with intake, policy gate, CAB approval, exception handling, evidence pack, and audit trail flow.
7. Use the new graph node/edge record import path for inventory-driven gallery scenarios such as identity/authentication flow, privileged access review, Kubernetes/service-mesh topology, and data-platform lineage maps, keeping generated connector ids stable for structural diffs.
