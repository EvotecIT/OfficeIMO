# OfficeIMO.Visio Visual Audit

Date: 2026-05-31
Branch/worktree: `codex/visio-premium-roadmap` at `C:\Support\GitHub\OfficeIMO-visio-premium-roadmap`

## Evidence Reviewed

- `dotnet build OfficeIMO.Examples\OfficeIMO.Examples.csproj -c Release --framework net8.0 --no-restore -nodeReuse:false -p:UseSharedCompilation=false`
- `OfficeIMO.Examples\bin\Release\net8.0\OfficeIMO.Examples.exe --visio-showcase --visio-preview`
- `OFFICEIMO_UPDATE_VISIO_PREMIUM_BASELINES=1 dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --framework net8.0 --filter "FullyQualifiedName~VisioPremiumVisualBaseline" --logger "console;verbosity=minimal" /p:NoWarn=CS8600%3BCS8602%3BCS8604`
- `dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --framework net8.0 --filter "FullyQualifiedName~VisioPremiumVisualBaseline" --logger "console;verbosity=minimal" /p:NoWarn=CS1591%3BCS8600%3BCS8602%3BCS8604`
- Generated summary: `OfficeIMO.Examples\bin\Release\net8.0\Documents\Visio Showcase\showcase-summary.md`
- Preview gallery: `OfficeIMO.Examples\bin\Release\net8.0\Documents\Visio Showcase\Preview\index.html`
- Approved premium baselines: `OfficeIMO.Tests\Visio\VisualBaselines\officeimo-visio-premium-*.png` and `.svg`

The latest local showcase proof generated 26 VSDX files and 52 PNG/SVG preview files through the optional Visio Desktop export path.

## What Is Actually Done

- The `OfficeIMO.Visio` package license text now matches the repository MIT license story.
- `--visio-showcase` is wired as the main showcase smoke path and fails if no VSDX files are generated.
- `--visio-showcase --visio-preview` exports reviewable PNG/SVG artifacts when Microsoft Visio is available.
- `--visio-premium` / `--premium-visio` generates a dedicated six-diagram premium subset.
- The six-diagram premium set now lives in reusable `OfficeIMO.Visio.VisioPremiumGallery`; the example is a thin caller.
- `VisioPremiumVisualBaselineTests` exports the six premium diagrams through Microsoft Visio desktop, stores approved PNG/SVG artifacts, compares regenerated previews against the approved baseline, and writes expected/actual/diff PNG artifacts with pixel statistics when a PNG preview changes.
- `CreateInspectionSnapshot()` now captures deterministic document/page/master/shape/connector/Shape Data/User cell structure and `VisioInspectionDiff` reports stable structural differences.
- `CreateStencilProfile()` now summarizes generated-master, package-backed, and basic-geometry shape usage plus Shape Data keys, semantic kind usage, stencil catalogs, categories, tags, and source package paths from inspection snapshots; generated and package-backed stencil provenance survives save/load.
- Diagram titles now use a readable title style instead of inheriting white text from filled emphasis shapes.
- Sequence self-message labels are now sized, kept outside the loop, and flipped left when the participant is near the page edge.
- Architecture regions, block regions, network zones, graph zones, and topology subnets now use separate header-style caption adornments instead of centered background text.
- Caption-aware layout reserves top clearance so titles/legends and zone headers do not collide, and the visual quality analyzer treats generated background captions as intentional adornments.
- Connector label cleanup now runs a second stabilization pass over all current labels and ignores generated adornment shapes, reducing order-sensitive label collisions in dense pages.
- Explicit obstacle-aware orthogonal routing can now reroute connectors around unrelated top-level shapes, and `PolishDiagram` can opt into that pass before connector-label cleanup.
- Premium enterprise, cloud, process, and dark-safe style presets now exist in the reusable theme catalog and the premium gallery uses them instead of relying only on the older generic presets.
- The first premium examples were tightened after rendered inspection:
  - visible titles on timeline, swimlane, sequence, graph, cloud, and network samples;
  - fewer obvious connector-label collisions in the premium examples;
  - simplified network and graph sample clutter where current automatic routing/zones hurt the look.
  - the premium incident sequence now uses the repaired self-message placement instead of avoiding the scenario.
  - premium cloud/network examples now show real zone headers again after the caption placement fix.
  - premium network and executive samples were adjusted after baseline inspection to avoid clipped/tight nodes and to show client access-switch connectivity.

## Bugs And Gaps Found

- Premium examples are credible smoke artifacts, not yet market-leading gallery material.
- Architecture and network zones are still visually large and generic, although captions now use header placement instead of center placement.
- Connector labels are improved but still not finished. Cleanup now revisits all current labels and ignores generated adornments, but it still does not understand lifeline bands, connector crossings, or zone-aware preferences deeply enough for every dense diagram.
- Sequence self-message placement was fixed for the reviewed long-label case, but sequence activations, notes, and richer interaction fragments are still missing.
- Simple connector-to-shape crossings are now covered by deterministic obstacle-aware routing, but dense network layouts still need zone-aware and crossing-aware route planning.
- Graph zones can overlap or dominate the diagram when groups span nodes across layers; the premium sample avoids zones for now.
- Stencil-backed visuals are not broad enough. Several examples still rely on basic geometry rather than recognizable first-party/native/external stencil symbols, although the richer stencil profile API can now measure catalog/category/source-pack coverage instead of relying on manual inspection.
- Imported master artwork children are filtered out of stencil profile counts, so package-backed profiles report logical placed stencil shapes instead of internal master artwork.
- Premium PNG baseline failures now include a rendered `.diff.png`, changed-pixel count, max channel delta, tolerance, and allowed-difference settings; SVGs still use canonicalized text comparison for Visio's unstable generated CSS class numbering.

## Recommended Next Steps

1. Extend obstacle-aware orthogonal routing into zone-aware, group-aware, and connector-crossing-aware route planning.
2. Continue connector label cleanup with lifeline/zone-aware preferences and connector-crossing avoidance.
3. Replace basic geometry in the premium gallery with native/external/first-party stencil-backed nodes where available.
4. Continue theme polish with diagram-specific spacing, title treatment, connector weights, and baseline-reviewed dark/print-safe defaults.
5. Add richer sequence-diagram features: activations, notes, combined fragments, and visible error/remediation bands.
6. Expand stencil extraction beyond current catalog/category/source-pack provenance into connection-point, icon/preview, package-family, and typed stencil-family profiles.
7. Use inspection diffs next to visual baseline failures so review output explains both structural and rendered changes.

## Status Call

This branch improves the current state and proves generation, preview export, approved premium visual baselines with PNG diff artifacts, structural inspection/diff snapshots, persisted catalog/category/source-pack stencil usage profiles, the first obstacle-aware routing pass, and the first premium theme preset slice end to end. It does not finish the full premium Visio goal. The remaining work is real product work in deeper layout, zone-aware routing, labels, broader stencils, theme polish, richer sequence semantics, deeper stencil metadata extraction, and richer structural-plus-rendered review artifacts.
