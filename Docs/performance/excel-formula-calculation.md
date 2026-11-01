# Excel formula calculation performance

## Workload and stop rule

The target is a fully calculated 200-product, 1,000-order report with 2,048
formulas: category lookup, numeric expressions, COUNTIFS/SUMIFS category totals
and grand totals. Every cached value is validated against totals computed from
the input rows. Compare exact INDEX/MATCH hits and missing products, VLOOKUP
missing products, and XLOOKUP missing products. Timing excludes workbook setup
and validation. Measurements stay opt-in, outside correctness CI.

First reduce repeated shared-string snapshot work. Accept a candidate only when
all outputs remain correct, allocation falls by at least 90% in the affected
lookup cases, and balanced measurements improve on both processor domains.
Stop after that owner-level change is qualified; additional indexing, parser
or aggregate optimizations need separate evidence rather than speculative work.

## Baseline: 1 November 2026

Source before optimization: `bace5c13c9646165bcfe0e97f49dfa457a0a5e35`.
The earlier lookup-fix candidate `655422df3d` measured roughly 13.0 GB allocated
for INDEX/MATCH hits, 13.1 GB for missing products, 6.7 GB for VLOOKUP, and
25.9 GB for XLOOKUP per report calculation. This is cumulative allocation,
not retained memory. The published 3.5.3 missing-product INDEX case fails
validation with 58 of 2,048 caches absent and has no valid comparison median.

The diagnostic host is Windows x64 on AMD Ryzen 9 9950X3D2 (16 cores/32 logical
processors), High Performance power plan, normal process priority. Native L3
topology reports masks `0xFFFF` and `0xFFFF0000`. Each complete cohort uses one
fixed domain, rotating engine order, one warmup and three measured iterations.
Background workstation activity and only three samples limit timing precision.
Keep domains separate and retain failed samples. PowerForge/PSPublishModule
3.0.170 owns timing, ordering, allocations and normalized reports.

## Attempts

- Accepted: reuse SharedStringCache within the existing formula calculation
  context. The editable-cell path previously rebuilt the workbook shared-string
  snapshot for every shared-string read. One disposable pass context now owns
  the snapshot across workbook, standalone-sheet and inspection operations.
  Referenced-sheet wrappers borrow it; later reads ignore a disposed context.
  Tracked input mutations invalidate the snapshot, and a new pass observes direct
  Open XML edits. Custom callbacks retain their documented deterministic,
  mutation-free contract. No public API or dependency is added.

## Qualified comparison: 1 November 2026

The [normalized evidence](excel-formula-calculation-2026-11-01.json) contains
the exact source and assembly identities, all 48 measured samples, per-domain
summaries, input dimensions and validation counts. The candidate is
`0db5cff3154a8678e636ad40dc1e68be46d9fd12`; the baseline already contains the
correctness fixes, so both sides perform the same complete calculation.

On mask `0xFFFF`, median times across the four cases fall from 3.25–12.52 seconds
to 0.15–0.30 seconds. On mask `0xFFFF0000`, they fall from 4.13–12.16 seconds
to 0.20–0.27 seconds. The per-case improvements are approximately 19–46 times.
Allocation falls by 97.7–98.4%: INDEX/MATCH uses about 252 MB instead of
13.0–13.1 GB, VLOOKUP about 155 MB instead of 6.7 GB, and XLOOKUP about
415 MB instead of 25.9 GB. These are cumulative managed allocations during
calculation, not peak or retained RAM.

Every measured operation produces and validates all 2,048 caches. Both domains
meet the stop rule, so this change adds no lookup indexing or aggregate cache.
The small sample count and shared workstation activity limit timing precision;
these measurements describe this report, host and runtime, rather than every
Excel formula or workbook size. Native Excel was unavailable on the measurement
host; no fresh native interoperability result is claimed by this performance run.

## Reproduction

The maintained [FormulaCalculationBenchmarks](../../Benchmarks/ExcelReaderTyped/FormulaCalculationBenchmarks.cs)
fixture exercises the same four report shapes in the existing opt-in benchmark
project. Its README documents running it and selecting saved baseline/current
assemblies. BenchmarkDotNet owns measurement for that suite; the captured
diagnostic comparison above uses PowerForge/PSPublishModule's shared runner.
Both exclude input construction and output validation from timing. Use fresh
workbooks, matching Excel/Core assemblies, equal placement and settings, and
separate output roots. Never treat an incomplete baseline as equivalent work.

## Lookup expansion comparison: 1 November 2026

The [expansion evidence](excel-formula-expansion-2026-11-01.json) compares the
complete shared-string-optimized source `67ab8dd54bcc17697df755b5b69bb36316056827`
with `ef7fc8843ef80862d250d97fc0fd3e3906072689`. It records the exact Excel/Core
assembly hashes, all 144 measured report operations and the complete summaries.
Both versions calculate the same four inputs and validate all 2,048 cached values.

The candidate adds lookup modes, reference/array INDEX and named expressions.
Lazy reference-backed lookup sources avoid eager materialization, per-pass name
catalogs avoid repeated definition scans, and literal lookup keys avoid wildcard
regex construction. Nine measured iterations per engine/case use rotated order
and each fixed processor domain, with one warmup. The final comparison includes
the shared blank-key correction as well as the literal-key regex optimization.
VLOOKUP timing remains mixed across domains, so this comparison establishes its
allocation reduction without claiming a consistent timing improvement.

| Report | Domain 0 median, baseline → candidate | Domain 1 median, baseline → candidate | Formula allocation, baseline → candidate |
| --- | --- | --- | --- |
| INDEX/MATCH hit | 236.95 → 148.70 ms | 191.74 → 115.64 ms | 251.6–252.0 → 150.3–151.1 MB |
| INDEX/MATCH missing product | 224.37 → 138.98 ms | 180.31 → 125.35 ms | 251.6–252.0 → 153.0–153.4 MB |
| VLOOKUP missing product | 144.96 → 127.83 ms | 115.11 → 122.44 ms | 154.6–155.0 → 148.0–148.4 MB |
| XLOOKUP missing product | 352.56 → 143.94 ms | 296.79 → 116.64 ms | 414.4–414.8 → 150.4–150.7 MB |

Allocation is cumulative managed allocation inside Calculate, expressed in decimal
MB; it is not peak or retained RAM. INDEX/MATCH allocation falls approximately
39–40%, XLOOKUP approximately 64%, and VLOOKUP approximately 4%. INDEX/MATCH medians
improve approximately 1.44–1.66 times and XLOOKUP 2.45–2.54 times against this already
optimized baseline. VLOOKUP is approximately 12% faster on domain 0 and 6% slower
on domain 1, with overlapping sample ranges and 13–18 ms standard deviations.
Those mixed timings do not establish a portable speedup or regression. Background
development activity remains uncontrolled; the recorded 12% CPU snapshot before
the final run is not a continuous trace.

## Binary lookup scale

The maintained [FormulaLookupScaleBenchmarks](../../Benchmarks/ExcelReaderTyped/FormulaLookupScaleBenchmarks.cs)
compares 256 linear and binary XLOOKUPs over identical sorted numeric sources at
512 and 8,192 keys. Setup and full result validation are outside timing. Each
processor domain ran the complete four-case matrix with BenchmarkDotNet 0.15.8
ShortRun, one launch, three warmups and three measured iterations at normal
priority. The evidence identifies the measured `a94fbfeb97ecfa7f0f1df41d0f500bd700fcc4a8`
assemblies; binary traversal is unchanged in the final report candidate.

At 8,192 keys, linear search means are 530–630 ms and binary means 110–120 ms.
Allocation is approximately 754 MiB for linear search and 11.5–11.6 MiB for binary
search. At 512 keys, allocation falls from approximately 55 MiB to 2.7–5.0 MiB,
but timing is variable and does not establish a consistent binary speedup.
All 256 cached results were checked after every measured operation. Binary search
requires sorted input; this numeric cohort does not qualify arbitrary mixed-type
or locale-dependent text ordering. BDN's three-sample confidence intervals are
wide, so these results demonstrate the larger-case scaling benefit without
claiming precise portable timing ratios.

All final report operations and scale cases pass output validation. No persistent
lookup index, parser cache, timing assertion or memory budget is added to ordinary
correctness CI. Native Excel was unavailable for a completed new oracle on the
measurement host; these performance runs establish equivalent validated work,
not native formula parity.
