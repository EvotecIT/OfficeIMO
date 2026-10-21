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

The starting source is the merged report-workflow and repeated-save baseline
`4543902a9d2715877dac912220ecdd804f2bdb4b`. Current timings have not yet been
captured. Existing historical results prioritize table-heavy HTML and repeated
shared-font text; profiling must establish the current expensive owner before
changing code. Record the actual framework, machine, host activity, inputs,
commands and source/binary hashes with the next observations.

| Candidate investigation | Required evidence | Disposition |
| --- | --- | --- |
| Repeated table measurement and style resolution | Current stage/profile evidence and unchanged complete-page fingerprints | Awaiting baseline |
| Repeated font analysis, shaping or diagnostic copies | Current allocation stacks and retained semantic/appearance controls | Awaiting baseline |

Stop expanding an optimization when it does not improve the selected equivalent
workload, weakens output or defensive limits, or needs an unrelated redesign.
Keep or revert each attempted change with its measured result and reason. Freeze
runtime source before the final comparison and package proof; a later runtime
change invalidates the affected measurements.
