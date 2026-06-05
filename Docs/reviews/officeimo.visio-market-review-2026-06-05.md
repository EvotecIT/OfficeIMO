# OfficeIMO.Visio Market Review

Date: 2026-06-05
Branch/worktree: `codex/visio-market-review-20260605` at `C:\Support\GitHub\OfficeIMO-visio-market-review-20260605`
Baseline: `origin/master` at `b31d0f80` (`Merge pull request #1884 from EvotecIT/codex/pdf-premium-next-20260604`)

## Proof Run

```powershell
dotnet build .\OfficeIMO.Visio\OfficeIMO.Visio.csproj -c Release --no-restore -m:1 -nr:false /clp:ErrorsOnly
dotnet build .\OfficeIMO.Examples\OfficeIMO.Examples.csproj -c Release --framework net8.0 -m:1 -nr:false /clp:ErrorsOnly
dotnet .\OfficeIMO.Examples\bin\Release\net8.0\OfficeIMO.Examples.dll --visio-showcase --visio-native-preview
.\.github\scripts\Assert-VisioShowcaseSummary.ps1 -ShowcasePath '.\OfficeIMO.Examples\bin\Release\net8.0\Documents\Visio Showcase' -RequirePreviewsPerDiagram -RequireProofsPerDiagram -RequireNativePreviewFormatsPerDiagram
dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --framework net8.0 --no-build --filter "FullyQualifiedName~Visio" --logger "console;verbosity=minimal"
```

Result: Visio and example builds passed; regenerated showcase validation passed with `44` diagrams, `44` packages, `88` previews, `132` proofs, `264` artifact hashes, complete preview/proof/evidence/stencil coverage, and `44/44` clean visual-quality proofs with `0` warning-level issues and `0` errors. The default Visio test filter passed `740/740`. A direct pre-gating run of `OfficeIMO.Tests.VisioPremiumVisualBaselineTests.PremiumGalleryPreviewsMatchApprovedBaselines` became inactive inside local Microsoft Visio COM export after several successful premium exports and hit the five-minute blame hang timeout, so this branch makes the desktop baseline lane explicit via `OFFICEIMO_RUN_VISIO_PREMIUM_DESKTOP_BASELINES=1` or `OFFICEIMO_REQUIRE_VISIO_PREMIUM_BASELINES=1`, and the smaller desktop validator smoke tests explicit via `OFFICEIMO_RUN_VISIO_DESKTOP_VALIDATION=1`, while leaving native/package/showcase proof in the default lane.

## Where We Are

OfficeIMO.Visio is already a serious product foundation, not a toy VSDX writer. Current `master` includes package creation, load/edit/save, preservation of unsupported content, validators, native comments, containers, layers, hyperlinks, Shape Data, User cells, protection, page settings, background pages, groups, masters, style themes, first-party stencils, package-backed stencil catalogs, native SVG/PNG preview export, optional desktop Visio validation/export, visual baselines, inspection snapshots, stencil profiles, and high-level builders for flowcharts, blocks, dependencies, architecture, networks, topology, swimlanes, org charts, timelines, sequences, and generic graphs.

The most important change in status is that the product is no longer missing the basic engine. The remaining gap is market execution: generated diagrams must look consistently premium, showcase proof must be easy to review in CI and PRs, public docs must match real APIs, and stencil/package workflows must feel first-class at scale.

## Issues Found In This Review

- Stale product proof metadata: the assessment and roadmap still pointed at the 2026-06-04 worktree and `731/731` test result. Updated to the current 2026-06-05 worktree and `738/738` proof.
- Stale package metadata: `OfficeIMO.Visio.csproj` still advertised `net50` / `net70` tags and a 2024 copyright. Updated tags to current target/product terms and copyright to 2026.
- Stale dependency wording: the Visio README described `System.IO.Packaging` as Windows-specific even though the package reference is not scoped to Windows. Updated the package README wording.
- Several Visio source files are now near or above the maintainability split signal: `VisioDocument.LoadCore.cs` at 919 lines, `VisioSwimlaneDiagramBuilder.cs` at 867, `VisioDuplicationExtensions.cs` at 857, `VisioOrgChartDiagramBuilder.cs` at 813, `VisioNetworkDiagramBuilder.cs` at 810, and `VisioGallery.cs` at 809. They are not broken, but future feature work should split by responsibility before adding much more behavior.
- The showcase proof was strong but still review-heavy. This branch adds versioned `showcase-summary.json` and `showcase-gallery.html` beside `showcase-summary.md`, with generated package, preview, review proof, artifact hash, `artifactCount`, top-level `proofTotals`, top-level `evidenceTotals`, diagram-level package/preview/proof/evidence metadata, structural shape/connector and Shape Data rollups, semantic-kind counts, stencil-backed/basic-geometry mix, connection-point coverage, parsed visual-quality `quality.*` counts, stencil catalog proof summaries parsed from `.stencil-profile.txt` artifacts, and stencil catalog coverage grouped by diagram. It also adds reusable `VisioShowcaseSummary` artifact validation before summary publication plus a focused GitHub Actions workflow that independently validates `schemaVersion`, summary counts/package-to-preview/proof grouping, top-level proof and evidence totals, native/desktop preview-format completeness, complete structural and complete review proof, shape/connector proof metrics, stencil backing metrics, connection-point counts, visual-quality proof schema and rollups, stencil proof summaries, stencil coverage, and recomputes file sizes/SHA-256 hashes through reusable scripts before uploading generated packages, native preview artifacts, review proof artifacts, and optional Microsoft Visio desktop preview artifacts for PR review. The current generated showcase reports `44/44` clean visual-quality proofs with `0` warning-level issues and `0` errors after adding builder polish, semantic background/adornment markers, and layout cleanup for the remaining showcase examples.

## What We Are Missing

### Market-Facing Output

The premium gallery is credible, but not yet the kind of gallery that ends the sales conversation by itself. We need more diagrams that look like real enterprise deliverables: cloud/security/data diagrams, Kubernetes/service-mesh diagrams, identity/auth flows, CI/CD topologies, incident/runbook sequences, rack/server/network maps, and executive dependency views. The library should show outcomes, not just APIs.

### Dense Layout And Label Strategy

Routing, label cleanup, obstacle avoidance, zone awareness, connector-path awareness, and sequence label placement have all improved. The remaining hard part is global dense-page minimization: crowded graphs, overlapping zones, multi-cluster dependencies, nested sequence fragments, and deep edge cases where connector labels compete with meaningful content.

### Stencil Platform Depth

First-party and package-backed stencils are real now. The next leap is deeper external master reuse, richer package-family metadata, unsupported imported master preservation, reusable stencil package export, and better documented native Visio stencil discovery. Users should be able to point OfficeIMO at a stencil library and quickly understand what is usable, what is approximate, and what needs fallback.

### Public Package Story

The README is now much closer to the truth, but website/product docs still need an API-accuracy pass. Public examples should prefer semantic builders, `VisioPremiumGallery`, package-backed catalogs, native preview APIs, and validation output over coordinate-heavy snippets.

### Release-Quality Proof Workflow

Tests prove a lot, but a market-leading Visio product also needs easy artifact proof:

- generated `.vsdx` package list;
- native SVG/PNG preview gallery;
- optional desktop Visio SVG/PNG/PDF proof;
- inspection, stencil-profile, and visual-quality summaries;
- machine-readable showcase summary through `showcase-summary.json`, including top-level proof/evidence totals and diagram records that pair packages with preview and review proof artifacts plus native/desktop preview flags, inspection/stencil-profile/visual-quality evidence flags, complete review proof, shape/connector metrics, Shape Data counts, semantic-kind counts, stencil backing metrics, connection-point counts, stencil catalog proof summaries, and catalog coverage;
- browsable showcase artifact gallery through `showcase-gallery.html`, including headline proof and evidence metrics, evidence coverage, diagram review cards that pair packages with previews, inspection/profile/visual-quality proof links, hash evidence, shape/connector metrics, Shape Data counts, stencil backing metrics, connection-point counts, stencil catalog chips, and a stencil coverage table;
- PR/CI artifact links.

## Recommended Next Steps

### P0: Make The Public Story Accurate And Reviewable

Update website/docs/examples to use current real APIs and premium scenarios. The new workflow wires `--visio-showcase --visio-native-preview` output into CI artifacts using `showcase-summary.json`, `showcase-gallery.html`, native previews, top-level proof and evidence totals, structural shape/connector and Shape Data metrics, stencil-backed/basic-geometry mix, connection-point coverage, and `Structural Proof` inspection/profile/visual-quality artifacts; it also requires native SVG, native PNG, and complete review proof evidence per diagram in the native lane and provides a manual `run_desktop_preview` lane for `[self-hosted, Windows, Visio]` runners with Microsoft Visio installed.

### P1: Promote A Larger Premium Gallery

Expand from the current eight baseline-approved premium diagrams toward a market-facing set of 12-16 diagrams. This branch starts that expansion with privileged-access review, data-platform lineage, hybrid network operations, and process governance review graphs covering request, policy, PAM, vault, target, audit, SIEM evidence, ingestion, quality, lake, warehouse, catalog, query API, analytics, edge/WAN, datacenter rack, compute host, storage, backup, operations monitoring, NOC review, CAB approval, exception handling, and evidence-pack flows. Continue prioritizing cloud/security/data/Kubernetes/identity/network/process/incident scenarios, and require every promoted diagram to pass package validation, native SVG/PNG preview, inspection snapshot, stencil profile, visual-quality proof, and visual baseline review.

### P2: Improve Dense Layout And Routing

Continue global route scoring, dense graph crossing minimization, connector-label placement, zone-aware layout, and nested sequence-fragment behavior. Add only contract-focused tests that prove real visible drift, structural drift, or user-facing layout regressions.

### P3: Harden Desktop Visio Proof Automation

Keep the dependency-free native preview lane as the default CI proof, but make the optional Microsoft Visio desktop lane more robust before treating it as always-on evidence. The current local run proved package/native/showcase quality, but `PremiumGalleryPreviewsMatchApprovedBaselines` became inactive inside COM export after several successful premium exports. Next work should isolate Visio desktop export batches, add per-document timeout/retry diagnostics, record the diagram/export where automation stalls, and clean up spawned Visio instances deterministically without touching a user's already-open Visio session.

### P4: Deepen Stencil Workflows

Preserve more package-backed master content, extend metadata extraction, document native pack discovery, add reusable stencil export, and keep migration presets conservative. This is the path to making stencils a platform rather than a catalog sidebar.

### P5: Keep Editing Existing Diagrams Safe

Grow loaded-diagram editing around real workflows: richer containers, swimlane maintenance, comment threading/author workflows, data graphics, Shape Data schema validation, selected-subset relayout, replace-master for external masters, and round-trip tests against more Visio-authored fixtures.

## Product Bet

The best market position is not "free Aspose clone" and not "raw Open XML helper." The winning position is:

> A dependency-light, server-safe C# Visio library where developers describe real diagrams, choose first-party or external stencils, get professional output, and can prove the result without installing Visio.

That means OfficeIMO.Visio should keep the VSDX core conservative and keep pushing reusable product value into `OfficeIMO.Visio` itself: builders, layouts, stencils, themes, renderers, validators, inspection, and proof artifacts. Examples, websites, and wrappers should stay thin.
