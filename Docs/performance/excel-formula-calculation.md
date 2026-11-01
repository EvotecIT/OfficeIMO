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
with `7c7cf4b0769028732b65d39898df6d0dc072d92f`. It records the exact Excel/Core
assembly hashes, all 144 measured report operations and the complete summaries.
Both versions calculate the same four inputs and validate all 2,048 cached values.

The measured lookup candidate adds lookup modes, reference/array INDEX and named expressions.
Lazy reference-backed lookup sources avoid eager materialization, per-pass name
catalogs avoid repeated definition scans, and literal lookup keys avoid wildcard
regex construction. Nine measured iterations per engine/case use rotated order
and each fixed processor domain, with one warmup at Normal priority. This
candidate includes the native-qualified blank/reference, source-error,
conditional and binary-duplicate corrections, plus typed callback preservation.

| Report | Domain 0 median, baseline → candidate | Domain 1 median, baseline → candidate | Formula allocation, baseline → candidate |
| --- | --- | --- | --- |
| INDEX/MATCH hit | 537.45 → 328.60 ms | 304.76 → 193.26 ms | 251.9–252.0 → 151.0–151.1 MB |
| INDEX/MATCH missing product | 438.83 → 272.97 ms | 313.41 → 206.31 ms | 251.9–252.0 → 153.3–153.3 MB |
| VLOOKUP missing product | 256.59 → 275.77 ms | 193.45 → 200.57 ms | 154.9–154.9 → 148.3–148.3 MB |
| XLOOKUP missing product | 549.47 → 289.48 ms | 435.39 → 176.99 ms | 414.7–414.8 → 150.6–150.7 MB |

Allocation is cumulative managed allocation inside Calculate, expressed in decimal
MB; it is not peak or retained RAM. INDEX/MATCH allocation falls approximately
39–40%, XLOOKUP approximately 64%, and VLOOKUP approximately 4%. INDEX/MATCH medians
correspond to 1.52–1.64 times faster calculation and XLOOKUP 1.90–2.46 times in
these runs against the already optimized baseline. VLOOKUP medians are about
4–7.5% higher, with overlapping sample ranges and 42–110 ms standard deviations.
Its timing effect remains inconclusive; the evidence establishes its allocation
reduction without claiming a speedup.

The shared development host was heavily loaded: the CPU snapshot before the
report comparison was 97%. This is not a continuous trace. Large sample dispersion
limits portable timing conclusions; both domains, every sample and slower results
remain in the evidence.

## Binary lookup scale

The maintained [FormulaLookupScaleBenchmarks](../../Benchmarks/ExcelReaderTyped/FormulaLookupScaleBenchmarks.cs)
compares 256 linear and binary XLOOKUPs over identical sorted numeric sources at
512 and 8,192 keys. Setup and full result validation are outside timing. Each
processor domain ran the complete four-case matrix with BenchmarkDotNet 0.15.8
ShortRun, one launch, three warmups and three measured iterations. The primary
cohorts explicitly apply Normal priority to all benchmark child processes and
measure the same `7c7cf4b0769028732b65d39898df6d0dc072d92f` lookup assemblies.

| Domain | Keys | Linear mean | Binary mean | Linear allocation | Binary allocation |
| --- | --- | --- | --- | --- | --- |
| 0 | 512 | 126.25 ms | 48.32 ms | 55.12 MiB | 4.91 MiB |
| 0 | 8,192 | 2004.74 ms | 346.45 ms | 753.45 MiB | 11.78 MiB |
| 1 | 512 | 236.28 ms | 99.35 ms | 55.12 MiB | 4.93 MiB |
| 1 | 8,192 | 1913.29 ms | 226.03 ms | 753.50 MiB | 11.82 MiB |

All 256 cached results were checked after every measured operation. Binary search
requires sorted input; this numeric cohort does not qualify arbitrary mixed-type
or locale-dependent text ordering. BDN's three-sample confidence intervals are
wide, so the table describes these runs rather than precise portable timing ratios.
The initial eight-case cohort used BDN's default High child priority; its complete
samples are retained separately in the same JSON. Keep processor and priority
cohorts separate when interpreting variation.

All final report operations and scale cases pass output validation. No persistent
lookup index, parser cache, timing assertion or memory budget is added to ordinary
correctness CI. The [native Excel lookup corpus](../../OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus/lookup-conformance.provenance.json)
separately qualifies 208 formulas, 232 typed result cells and 12 spill ranges after
save/reopen. The timed engine in both report lanes is OfficeIMO; desktop Excel
produces the independent correctness fixture.

The later `SUMPRODUCT` coercion correction affects a separate function that these
lookup workloads do not call. The recorded hashes identify the measured lookup
candidate; its timed lookup paths are unchanged by that correction.
