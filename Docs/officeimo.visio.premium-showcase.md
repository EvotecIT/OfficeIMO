# OfficeIMO.Visio Premium Showcase Plan

Date: 2026-05-31

## Goal

The showcase should prove that OfficeIMO.Visio can generate diagrams that are structurally valid, easy to edit in Visio, and visually credible without hand placement.

## Proof Levels

1. **Package proof**
   - Run `--visio-showcase`.
   - Validate every generated `.vsdx` with `VisioValidator`.
   - Write `showcase-summary.md` with the generated package list.

2. **Desktop proof**
   - Run `--visio-showcase --visio-preview` on a machine with Microsoft Visio.
   - Open, round-trip, and export PNG/SVG previews through `VisioDesktopValidator`.
   - Add preview files and gallery output as review artifacts.

3. **Visual baseline proof**
   - Store approved PNG/SVG previews for the core premium gallery under `OfficeIMO.Tests/Visio/VisualBaselines`.
   - Compare regenerated previews with `VisioPremiumVisualBaselineTests`.
   - Refresh approved artifacts deliberately with `OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES=1`.
   - Require Visio desktop for this lane with `OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES=1`; otherwise the test skips on machines without Visio.
   - On PNG drift, write expected, actual, and `.diff.png` artifacts with changed-pixel count and max channel delta.
   - Use `OFFICEIMO_VISIO_PREMIUM_BASELINE_PIXEL_TOLERANCE` and `OFFICEIMO_VISIO_PREMIUM_BASELINE_ALLOWED_DIFF_PIXELS` only for deliberate renderer-noise tolerance.
   - Treat structural validity and visual drift as separate gates.

## Premium Gallery Targets

- **Cloud architecture:** zones, trust boundaries, gateway, eventing, data store, runbook links, Shape Data.
- **Network segmentation:** internet, firewall, core switch, server zone, client zone, monitoring, labeled links.
- **Executive dependency graph:** layered service graph with owners, critical path, external dependencies, and grouped runtime zone.
- **Incident sequence:** actor/system/service participants, activations, notes, error path, and remediation path.
- **Release timeline:** milestones, risks, spans, decisions, callouts, and status metadata.
- **Swimlane process:** roles, phases, handoffs, exception path, compliance callouts, and readable lane headers.

## Design Bar

- Use semantic builders first; drop to low-level shapes only for deliberate visual accents.
- Prefer real first-party or package-backed stencil nodes where the domain benefits from recognisable symbols.
- Use stable IDs, hyperlinks, and Shape Data so diagrams are searchable and regeneration-friendly.
- Use the premium enterprise, cloud, process, and dark-safe presets as the default gallery palette, then keep typography, spacing, captions, and connector labels consistent across the gallery.
- Make every diagram explain a real business or operational scenario, not just API coverage.

## Next Implementation Slice

1. Extend the new obstacle-aware routing into zone/group-aware routes and add lifeline/zone-aware connector-label cleanup in dense premium diagrams.
2. Replace the remaining basic geometry with stencil-backed premium symbols where the domain supports it.
3. Continue baseline review of premium theme presets and promote only baseline-approved diagrams into website documentation and product screenshots.
4. Extend stencil profile reporting with richer catalog/category/source-pack metadata as external metadata extraction improves.
5. Broaden the gallery from six baseline-approved diagrams toward a larger market-facing set.
