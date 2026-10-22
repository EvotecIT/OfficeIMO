# HTML-to-PDF report performance

The [22 October evidence](../../Build/Project/Evidence/2026-10-22/pdf-report-layout-performance.json)
records report-layout checks and measurements of the shared HTML/PDF owners.
[Open work](../ROADMAP.md#release-wide-quality) remains in the product roadmap.

## Reproduce the measurements

Use the maintained BenchmarkDotNet projects and PowerForge comparison entrypoint:

```powershell
./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfhtmlpayload -RunMode full -Framework net10.0 -OutputRoot <task-output> -BuildArtifactsPath <task-build>
```

The HTML benchmark project's `HtmlPagedPurchaseTableBenchmarks` separates prepared
layout from PDF export at 250 and 2,500 rows. The export starts with parsed HTML
and writes to `Stream.Null`; it excludes parsing and filesystem I/O. The shared
payload benchmark parses HTML and returns PDF bytes. Keep these contracts
separate when interpreting results.

Run the purchase-table cases through that project's maintained BenchmarkDotNet
entrypoint, using separate output directories for each baseline and candidate:

```powershell
$env:DOTNET_PROCESSOR_COUNT = '3'
dotnet build ./OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net10.0 --artifacts-path <task-build> -m:3 -nodeReuse:false
dotnet <task-build>/bin/OfficeIMO.Html.Benchmarks/release_net10.0/OfficeIMO.Html.Benchmarks.dll --filter '*HtmlPagedPurchaseTableBenchmarks*' --inProcess --warmupCount 6 --iterationCount 12 --outliers DontRemove --exporters json --artifacts <task-output>/purchase
```

Replace the output placeholders before running. These are the filter and job
settings used for the purchase-table measurements below; the shared-payload
command above selects a different benchmark class.

The paired runs use the same inputs, fonts, page settings and benchmark drivers.
Complete content, saved-page geometry and independent rendering checks run
outside the timed section. Timing and allocation measurements stay opt-in.

## Final paired results: 22 October 2026

Baseline: `0612ee6a8bd8a18ba13eb7c3b64fcb104d182945`.
Candidate runtime: `765057d2bdf16706be5e0cdec228fd541abf8ee8`.
The Apple M4 host has ten logical processors and runs .NET 10.0.13. Each measured
launch sets `DOTNET_PROCESSOR_COUNT=3`, overriding the processor count reported
to .NET rather than pinning execution to three cores. The evidence retains both
the hardware count and the child-process environment. Runs use normal development
activity and OS scheduling. This is not a quiet-host, cold-start,
peak-memory or portable ranking study. Each run has six warmups, twelve measured
iterations and no outlier removal. Allocation is total managed allocation per
operation, not retained or peak memory.

The timing table describes the named runtime revision. The dated evidence also
records a separate check with later shared PDF text-rendering changes. That
combined-output check does not extend these timings, private-package results
or NativeAOT observations to the later combination.

| Purchase-table workload | Baseline mean | Candidate mean | Baseline allocation | Candidate allocation |
| --- | ---: | ---: | ---: | ---: |
| 250 rows, prepared layout | 185.50 ms | 167.73 ms | 220.74 MiB | 187.04 MiB |
| 250 rows, PDF export | 224.50 ms | 177.29 ms | 325.80 MiB | 283.22 MiB |
| 2,500 rows, prepared layout | 1,961.32 ms | 1,833.98 ms | 2,716.72 MiB | 2,382.22 MiB |
| 2,500 rows, PDF export | 6,007.78 ms | 2,558.61 ms | 4,528.64 MiB | 3,354.34 MiB |

The selected 2,500-row export mean is 57.4% lower, with 25.9% fewer allocated
bytes. Both public purchase-report PDFs are byte-identical to the baseline:
42 and 417 pages, with every row marker present once. The 250-row layout runs
have roughly 46 ms standard deviations, so their mean difference is not a
reliable layout-speed claim.

Two shared-payload pairs ran in baseline/candidate/candidate/baseline order:

| Payload | Baseline means | Candidate means | Baseline allocation | Candidate allocation |
| --- | ---: | ---: | ---: | ---: |
| Plain text | 38.14–40.99 ms | 39.45–39.46 ms | 96.14–96.16 MiB | 93.81–93.97 MiB |
| Table | 117.21–118.25 ms | 104.95–107.06 ms | 222.31–222.45 MiB | 199.68–199.69 MiB |
| Multilingual | 33.29–34.45 ms | 33.68–33.75 ms | 96.98–96.99 MiB | 94.62–94.79 MiB |

All three OfficeIMO payload PDFs are byte-identical to the baseline. The table
improvement repeats in both pairs. Plain-text and multilingual timings overlap
or change direction; no throughput improvement is claimed for them. The separate
nine-case shared comparison retains all successful engines and their differing
outputs; these results do not establish a speed lead over other libraries.

## Investigations and disposition

An instrumented 30-conversion table profile retained a byte-identical PDF. It
identified repeated CSS length setup/default parsing and repeated all-page scans
while enumerating tagged-PDF children. Sampling guided the following bounded
changes; it does not attribute the final gain to one change in isolation.

| Attempt | Observation and disposition |
| --- | --- |
| Reject empty, `auto`, `none` and `normal` length defaults early | First combined draft reduced table allocation by about 6.4%, but measured slower. Kept for lower allocation; no isolated throughput claim. |
| Cache successful short parsed length expressions | Clean `6a741e7f` trial: table 117.36–118.49 ms versus 116.85–119.42 ms; 203.83–204.26 MiB versus 222.33–222.45 MiB. Kept for lower allocation, not isolated speed. Cache is per resolver, at most 128 expressions of 128 characters; values resolve afresh for each context. |
| Reuse the private synchronous resolution workspace | Trial: 118.19–122.04 ms versus 117.68–118.13 ms; 200.28–200.57 MiB versus 222.45 MiB. Kept for allocation; all context inputs are overwritten and context controls pass. |
| Index tagged-PDF children once per write | Combined development trial: table 107.06–109.19 ms, about 199.6 MiB; 2,500-row export 6,129.1 to 2,572.7 ms with identical output. Kept after ordering, parent-identity and tagged-artifact checks; final results above replace this trial for current claims. |
| Integrated `ddddaa1c` checkpoint | Earlier paired export: 6,088.99 to 2,569.03 ms. The 250-row layout mean rose from 178.75 to 194.90 ms. Retained as historical evidence; later whitespace correction required the fresh candidate measurements above. |
| Preformatted-text correction | Nested styled text previously lost repeated spaces and line breaks. Saved RFC footer columns now align closely with the independent reference. This is a layout fix, not an optimization; the final candidate was remeasured. |
| Repeat glyph-outline/contour caching | Not attempted: the current profile points elsewhere. Earlier changes already addressed that owner. |

The [20 October evidence](../../Build/Project/Evidence/2026-10-20/pdf-workflow-reliability.json)
keeps its original `83efd` source and platform. Its 66–87% selected mean reductions
and 60–83% allocation reductions remain historical; they are not measurements of
this candidate or later shared-glyph source.

## Limits and next work

Generated table text, printed link destinations, ordinary scroll-panel pagination
and nested preformatted text have saved-file controls. The frozen captured reports
still have named font/resource and layout differences. The RFC case retains all
3,600 lexical tokens over 13 pages and fixes footer-column spacing, but its vertical
spacing differs from the native reference. Successful conversion and syntax checks
do not establish full appearance or accessibility equivalence.

Next, use the roadmap's licensed, font-matched corpus to qualify wide tables,
rotated labels and all-page native/assistive-technology behavior. Reprofile only
when those supported workloads identify another material cost. Stop an optimization
that weakens content, diagnostics, cancellation or resource limits, or requires an
unrelated redesign. Preserve slower and failed observations; do not repeat exhausted
ideas unless the relevant source has changed.


## Font, style and geometry reuse investigation: 22 October 2026

The baseline for this round is `8cc38e17423cc434a8306b1cd9cfc184ae30ba6c`.
The selected goal is lower warm conversion time and managed allocation on the
maintained 21 KiB payloads and purchase-table exports, with byte-identical saved
PDFs, unchanged diagnostic behavior and existing resource/cancellation limits.
The immutable baseline archive and compiled runtimes are retained with the run records.

The normal Apple M4 development host uses the repository SDK and .NET 10 runtime,
with `DOTNET_PROCESSOR_COUNT=3`, no affinity or power changes, six warmups,
twelve measured iterations and no outlier removal for final paired measurements.
Use the maintained commands above. Keep different workloads and earlier source
revisions in separate cohorts; allocation is total managed allocation, not peak
or retained memory.

| Attempt | Observation and disposition |
| --- | --- |
| Reuse short built-in fallback measurements within one prepared conversion | Initial draft; full text/font/size/face key, font-definition snapshot invalidation, 1,024 entries and 256-character text/family limits. Host-selected faces and caller shaping providers keep their ordinary path. Included in the initial paired trial below; final qualification is pending. |
| Reuse resolved border and outline values | Initial draft; one resolver, immutable computed-style identity and every original numeric/color/character context. Resolved parent boxes are never memoized. Included in the initial paired trial below; final qualification is pending. |
| Reuse private glyph-placement workspace | Owner-confirmed separate placement workspace draft. Immutable origin outlines remain unchanged; every point is copied before the original placement arithmetic, with bounded writer-local storage. First compiled trial below; final qualification is pending. |

Stop or revert an attempt that changes saved content, geometry, diagnostics,
requested compliance or bounded-resource behavior. Keep failed and slower trials.
A smaller allocation without a reliable throughput improvement is reported as
such, and does not close the speed objective. No public API, provider or production
dependency change is part of this round.

The initial font/stroke draft passed six focused .NET 10 cases and all nine
payload admissions, with the three OfficeIMO PDFs byte-identical to the fresh
baseline. One preliminary pair used the maintained in-process benchmark job
with six warmups and twelve measurements:

| Payload | Baseline mean (ms) | Draft mean (ms) | Baseline allocated (MiB) | Draft allocated (MiB) |
| --- | ---: | ---: | ---: | ---: |
| Plain | 45.61 | 34.12 | 93.86 | 58.39 |
| Table | 118.39 | 101.49 | 199.89 | 162.35 |
| Multilingual | 35.43 | 27.86 | 94.67 | 58.26 |

This is an initial dirty-source trial, before geometry reuse, rather than final
qualification or an isolated attribution to either change. All raw samples,
actual runtime inputs and the initial test-fixture compilation failure are
retained. The fixture was corrected to use the existing linked scalar-advance
font helper; no product behavior was changed to satisfy that failed compile.

The next draft included the private glyph-placement workspace. It passed 114
focused PDF cases and the six HTML cases on .NET 10, all nine payload admissions
and syntax checks, and byte-identical OfficeIMO output for all three payloads.
The first alternating-glyph assertion failed on both the draft and immutable
baseline because separate documents repeat their own document-wide notices.
The corrected comparison retains every distinct warning with its exact bounds,
without multiplying those notices. Both failed controls remain in the raw record.

| Payload | Baseline mean (ms) | Geometry draft mean (ms) | Draft standard deviation (ms) | Draft allocated (MiB) |
| --- | ---: | ---: | ---: | ---: |
| Plain | 45.61 | 45.98 | 11.74 | 47.25 |
| Table | 118.39 | 100.01 | 5.20 | 154.33 |
| Multilingual | 35.43 | 27.43 | 0.49 | 50.73 |

The plain-text mean lost the earlier trial's timing benefit and varied widely;
its samples are retained rather than removed. Allocation fell for all three
payloads. This preliminary trial does not close the speed objective. Final
qualification uses a stable committed candidate, rotated baseline/candidate
pairs, larger purchase reports and saved warning/geometry controls.
