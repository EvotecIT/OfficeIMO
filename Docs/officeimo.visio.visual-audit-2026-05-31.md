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
- Approved premium baselines: `OfficeIMO.Tests\Visio\VisualBaselines\officeimo-visio-premium-*.png`, `.svg`, `.inspection.txt`, and `.stencil-profile.txt`

The latest local showcase proof generated 26 VSDX files and 52 PNG/SVG preview files through the optional Visio Desktop export path.

## What Is Actually Done

- The `OfficeIMO.Visio` package license text now matches the repository MIT license story.
- `--visio-showcase` is wired as the main showcase smoke path and fails if no VSDX files are generated.
- `--visio-showcase --visio-preview` exports reviewable PNG/SVG artifacts when Microsoft Visio is available.
- `--visio-premium` / `--premium-visio` generates a dedicated eight-diagram premium subset.
- The eight-diagram premium set now lives in reusable `OfficeIMO.Visio.VisioPremiumGallery`; the example is a thin caller.
- `VisioPremiumVisualBaselineTests` exports the eight premium diagrams through Microsoft Visio desktop, stores approved PNG/SVG artifacts, compares regenerated previews against the approved baseline, and writes expected/actual/diff PNG artifacts with pixel statistics when a PNG preview changes.
- The same premium baseline lane now stores approved inspection snapshots and stencil profiles for every premium diagram, supports structural-only refresh with `OFFICEIMO_UPDATE_VISIO_PREMIUM_STRUCTURAL_BASELINES=1`, and writes expected/actual/`.diff.txt` artifacts when structural or stencil-profile output drifts.
- `CreateInspectionSnapshot()` now captures deterministic document/page/master/shape/connector/Shape Data/User cell structure and `VisioInspectionDiff` reports stable structural differences.
- `CreateStencilProfile()` now summarizes generated-master, package-backed, basic-geometry, and stencil-backed shape usage plus Shape Data keys, semantic kind usage, stencil catalogs, categories, tags, source package paths, and connection-point richness from inspection snapshots; generated and package-backed stencil provenance survives save/load.
- Graph, architecture, network, flowchart, block-diagram, swimlane, timeline, and sequence stencil nodes selected from a catalog now keep their catalog name in inspection/profile output and built-in catalog stencil nodes inherit theme styling, so the premium executive dependency graph, cloud architecture gallery, network segmentation gallery, print audit trail gallery, technical topology gallery, governed process gallery, release timeline gallery, and incident sequence gallery can use first-party stencils without losing the polished rendered look.
- Diagram titles now use a readable title style instead of inheriting white text from filled emphasis shapes.
- Sequence self-message labels are now sized, kept outside the loop, and flipped left when the participant is near the page edge.
- Sequence activations now exist as first-party sequence-stencil shapes with semantic profile metadata, dedicated layers, and premium incident-sequence gallery coverage.
- Sequence combined fragments now exist as first-party sequence-stencil frames with semantic profile metadata, dedicated layers, generated top-left labels, visual-quality exemptions for intentional frame crossings, and premium incident-sequence gallery coverage.
- Sequence notes now use collision-aware candidate placement against existing shapes, sequence connector labels, and connector segments, persist requested/resolved placement metadata, and have premium incident-sequence gallery coverage.
- Architecture regions, block regions, network zones, graph zones, and topology subnets now use separate header-style caption adornments instead of centered background text.
- Caption-aware layout reserves top clearance so titles/legends and zone headers do not collide, and the visual quality analyzer treats generated background captions as intentional adornments.
- Connector label cleanup now runs a second stabilization pass over all current labels and ignores generated adornment shapes, reducing order-sensitive label collisions in dense pages.
- Explicit obstacle-aware orthogonal routing can now reroute connectors around unrelated top-level shapes, optional zones/containers, and existing connector paths, and `PolishDiagram` can opt into those passes before connector-label cleanup.
- Premium enterprise, technical, cloud, process, print-safe, and dark-safe style presets now exist in the reusable theme catalog, and the current premium gallery uses the baseline-approved preset set instead of relying only on older generic presets.
- The dark-safe preset now uses readable connector-label text on white exported pages while keeping high-contrast connector lines and dark filled shapes.
- The premium gallery now includes rendered technical topology and print audit trail scenarios, so the technical and print-safe theme presets have PNG/SVG baseline proof instead of only package-level smoke coverage.
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
- Connector labels are improved but still not finished. Cleanup now revisits all current labels, ignores generated adornments, and can prefer common endpoint zones, but it still does not understand lifeline bands deeply enough for every dense diagram.
- Sequence self-message placement was fixed for the reviewed long-label case, semantic notes now exist as first-party sequence-stencil shapes with collision-aware placement, and activation bars plus combined fragment frames now have first-party stencil/profile coverage. Richer guarded/partitioned interaction fragments are still missing.
- Simple connector-to-shape crossings and connector-to-connector crossings are now covered by deterministic obstacle-aware route scoring, but dense network layouts still need group-aware and deeper whole-page route minimization.
- Graph zones can overlap or dominate the diagram when groups span nodes across layers; the premium sample avoids zones for now.
- Stencil-backed visuals are not broad enough. All eight current premium baselines now carry first-party stencil catalog provenance for their domain shapes, including Sequence Diagram provenance for incident participants, activations, notes, and combined fragments, but deeper sequence constructs still need guarded/partitioned fragments before the sequence gallery can be called fully premium.
- Generated stencil master instances now use renderer-friendly page references: generated first-party stencil shapes keep `Master` and local style deltas but no longer force `MasterShape="1"`, matching Visio-authored simple master instances. The gallery still needs a refreshed PNG/SVG baseline that promotes this path into a market-facing graph scenario.
- Imported master artwork children are filtered out of stencil profile counts, so package-backed profiles report logical placed stencil shapes instead of internal master artwork.
- Premium PNG baseline failures now include a rendered `.diff.png`, changed-pixel count, max channel delta, tolerance, allowed-difference settings, and the matching inspection/profile expected, actual, and `.diff.txt` context; SVGs still use canonicalized text comparison for Visio's unstable generated CSS class numbering.

## Recommended Next Steps

1. Extend obstacle-aware orthogonal routing from the current zone/container/crossing-aware options into group-aware and deeper whole-page route minimization.
2. Continue connector label cleanup with lifeline-aware preferences and denser premium-diagram label strategies.
3. Replace more basic geometry in the premium gallery with native/external/first-party stencil-backed nodes where available, while preserving rendered PNG/SVG quality.
4. Add more theme-specific scenarios beyond the current technical topology and print audit trail, especially cloud/security/data diagrams with stencil-backed nodes.
5. Add richer sequence-diagram features: guarded/partitioned fragments and visible error/remediation bands.
6. Promote generated-master stencil usage into a premium gallery baseline now that generated stencil instances use renderer-friendly master references.
7. Expand stencil extraction beyond current catalog/category/source-pack provenance into connection-point, icon/preview, package-family, and typed stencil-family profiles.
8. Use the new inspection/profile baseline artifacts to decide whether each visual drift is a rendering-only change, a shape/layout regression, or a stencil/profile regression.

## Status Call

This branch improves the current state and proves generation, preview export, approved premium visual baselines with PNG diff artifacts, approved inspection/profile baselines with text diff artifacts, structural inspection/diff snapshots, persisted catalog/category/source-pack stencil usage profiles, obstacle-aware routing with zone/container/crossing-aware options, zone-aware connector label cleanup, first-party sequence activations/fragments/notes, and the six-preset premium theme catalog with technical/print rendered baseline proof. It does not finish the full premium Visio goal. The remaining work is real product work in deeper layout, group-aware routing, lifeline-aware labels, broader stencils, richer sequence semantics, and deeper stencil metadata extraction.
