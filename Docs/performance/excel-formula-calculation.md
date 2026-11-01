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
with `8f1d519761a3f831ddafa15667ffb02d91944198`. It records the exact Excel/Core assembly hashes,
all 144 measured report operations and the complete summaries. Both versions
calculate the same four inputs and validate all 2,048 cached values.

The candidate adds lookup modes, reference/array INDEX and named expressions.
Lazy reference-backed lookup sources avoid eager materialization, per-pass name
catalogs avoid repeated definition scans, and literal lookup keys avoid wildcard
regex construction. Nine measured iterations per engine/case use rotated order
and each fixed processor domain, with one warmup at Normal priority. This final
measurement includes the native-qualified blank/reference, source-error,
conditional, binary-duplicate and typed-callback corrections, plus the shared
source adapter used by the qualified SUMPRODUCT input handling.

| Report | Domain 0 median, baseline → candidate | Domain 1 median, baseline → candidate | Formula allocation, baseline → candidate |
| --- | --- | --- | --- |
| INDEX/MATCH hit | 353.05 → 212.51 ms | 267.65 → 224.03 ms | 251.8–251.9 → 151.0–151.0 MB |
| INDEX/MATCH missing product | 319.54 → 225.35 ms | 239.45 → 171.78 ms | 251.8–251.9 → 153.3–153.3 MB |
| VLOOKUP missing product | 355.51 → 232.14 ms | 161.38 → 164.72 ms | 154.9–154.9 → 148.3–148.3 MB |
| XLOOKUP missing product | 639.25 → 204.21 ms | 406.92 → 163.12 ms | 414.7–414.7 → 150.6–150.7 MB |

Allocation is cumulative managed allocation inside Calculate, expressed in decimal
MB; it is not peak or retained RAM. The observed baseline/candidate median ratios
range from 1.19 to 1.66 for INDEX/MATCH and
2.49 to 3.13 for XLOOKUP.
VLOOKUP ratios are 1.53 on domain 0 and
0.98 on domain 1; a ratio above one means a smaller
candidate median. Keep these observations separate from portable speed guarantees.
Full sample ranges, standard deviations and slower results remain in the evidence.

The shared development host's CPU snapshot before this final report comparison
was 100%. It is not a continuous trace. Background development activity,
processor placement and finite sample counts limit timing precision.

## Binary lookup scale

The maintained [FormulaLookupScaleBenchmarks](../../Benchmarks/ExcelReaderTyped/FormulaLookupScaleBenchmarks.cs)
compares 256 linear and binary XLOOKUPs over identical sorted numeric sources at
512 and 8,192 keys. Setup and full result validation are outside timing. Each
processor domain ran the complete four-case matrix with BenchmarkDotNet 0.15.8
ShortRun, one launch, three warmups and three measured iterations. Normal priority
is explicitly applied to all benchmark children. These primary cohorts measure
the same `8f1d519761a3f831ddafa15667ffb02d91944198` assemblies as the report comparison.

| Domain | Keys | Linear mean | Binary mean | Linear allocation | Binary allocation |
| --- | --- | --- | --- | --- | --- |
| 0 | 512 | 79.39 ms | 86.53 ms | 55.12 MiB | 4.87 MiB |
| 0 | 8,192 | 1887.08 ms | 565.14 ms | 753.50 MiB | 11.78 MiB |
| 1 | 512 | 222.80 ms | 44.24 ms | 55.18 MiB | 5.06 MiB |
| 1 | 8,192 | 723.38 ms | 136.90 ms | 753.50 MiB | 11.86 MiB |

All 256 cached results were checked after every measured operation. Binary search
requires sorted input; this numeric cohort does not qualify arbitrary mixed-type
or locale-dependent text ordering. BDN's three-sample confidence intervals are
wide, so the table describes these runs rather than precise portable timing ratios.
The initial High-priority scale cohort from source `7c7cf4b` remains separately
identified in the JSON with its original source hashes. Keep source, priority
and processor cohorts separate when interpreting variation.

All final report operations and scale cases pass output validation. No persistent
lookup index, parser cache, timing assertion or memory budget is added to ordinary
correctness CI. The [native Excel lookup corpus](../../OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus/lookup-conformance.provenance.json)
separately qualifies 218 formulas, 242 typed result cells and 12 spill ranges after
save/reopen. The timed engine in both report lanes is OfficeIMO; desktop Excel
produces the independent correctness fixture.
