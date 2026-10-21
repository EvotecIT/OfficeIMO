# HTML-to-PDF report performance

This ledger records measurements and attempted improvements to the shared HTML
layout and PDF owners. It is not a capability matrix or a second roadmap.
[Open product outcomes](../ROADMAP.md#release-wide-quality) remain in the roadmap.

## Measurement contract

Measure complete HTML-to-PDF conversion separately from prepared layout and PDF
serialization. Use the same captured HTML, font bytes, page profile and requested
tags for every baseline/candidate pair. Validate complete content and saved-page
geometry outside timing; inspect rendered pages against independent references.
Keep layout corrections separate from byte-preserving optimizations when comparing
artifacts. Retain failed, slower and noisy observations.

Use the existing BenchmarkDotNet projects and PowerForge comparison entrypoint:

```powershell
./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfhtmlpayload -RunMode full -Framework net10.0 -OutputRoot <task-output> -BuildArtifactsPath <task-build>
```

The HTML benchmark project also separates parsing, styling, prepared layout,
Drawing projection and encoded output. Its purchase-table and page-fingerprint
lanes provide scaling and all-page controls. Timing and allocation measurements
remain opt-in and do not become ordinary correctness tests.

## Existing evidence and exhausted work

The [20 October qualification](../../Build/Project/Evidence/2026-10-20/pdf-workflow-reliability.json)
retains the original measured source and platform. On that Apple M4/.NET 10
source, eliminating repeated ink-contour work and adding a bounded glyph-outline
cache reduced selected warm HTML conversion means by 66–87% and allocation volume
by 60–83%. These are improvements over the earlier OfficeIMO baseline, not a
portable ranking or measurements of later integrated source. Later shared-glyph
changes require a fresh baseline before further claims.

Do not repeat the same contour/cache work without showing which underlying
source changed. Diagnostic accuracy, cancellation, font ownership and cache
bounds remain part of the contract; disabling them is not an optimization.

## Current investigation

The starting runtime is preparation commit
`0612ee6a8bd8a18ba13eb7c3b64fcb104d182945`, based on the merged report-workflow
and repeated-save source `4543902a9d2715877dac912220ecdd804f2bdb4b`.
Its four frozen print-report conversions and all nine shared-payload admission
cases succeed. The larger reports still retain documented resource and rendering
losses; successful conversion is not an appearance acceptance result.

An allocation and CPU profile of 30 conversions of the maintained 21 KiB table
payload identifies repeated CSS length contexts and parsing of nonnumeric
defaults as a useful first target. Tagged-PDF child enumeration also repeatedly
scans all page elements. The instrumented PDF is byte-identical to the unprofiled
admission artifact. Sampling identifies work to investigate; it does not establish
a throughput improvement or a portable memory budget.

| Candidate investigation | Required evidence | Disposition |
| --- | --- | --- |
| Reject empty and nonnumeric CSS length defaults before allocating a resolution context | Existing CSS math/context controls, equivalent saved payloads, and paired BenchmarkDotNet runs | About 6.4% less table allocation in the first combined draft; no throughput win established |
| Reuse parsed short length expressions within one style resolver | Recompute used values for each percentage/font/container context, retain limits and diagnostics, and compare complete payloads and table stages | Trial; implementation and focused context proof in progress |
| Index tagged-PDF children within one write | Preserve local and cross-page parent identity, child ordering and saved tagged artifacts; measure scaling | Profiled lead; coordinate shared-owner changes first |
| Repeated font analysis, shaping or diagnostic copies | Current allocation stacks and retained semantic/appearance controls | Lower priority in the current profile; do not repeat the earlier outline-cache work |

The first trial combines the generated table-content correction with early
rejection of empty, `auto`, `none` and `normal` lengths. Two paired launches in
baseline/candidate/candidate/baseline order use the unchanged benchmark, six
warmups, twelve measured iterations and no outlier removal. On the local Apple
M4/.NET 10 development host, table means are 118.95–120.86 ms for the baseline
and 122.28–125.22 ms for the draft; allocations are approximately 222.3 MiB and
208.2 MiB respectively. This is an allocation observation, not an accepted
speed improvement or a measurement of the guard in isolation. All three shared
OfficeIMO payload PDFs are byte-identical, and the 42-page/417-page purchase
controls retain every page SVG hash, logical text hash and diagnostic.

The next trial caches only successful immutable parse results, with at most
128 expressions of at most 128 characters per layout resolver. The existing
parser remains the grammar and limit owner. Percentage, font, viewport and
container values are resolved afresh for each use; neither resolved lengths nor
cross-document state are cached. Longer or additional expressions use the usual
uncached path.

Stop expanding an optimization when it does not improve the selected equivalent
workload, weakens output or defensive limits, or needs an unrelated redesign.
Keep or revert each attempted change with its measured result and reason. Freeze
runtime source before the final comparison and package proof; a later runtime
change invalidates the affected measurements.
