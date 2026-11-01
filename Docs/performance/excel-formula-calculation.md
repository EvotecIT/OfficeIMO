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

## Lookup expansion comparison

The next comparison uses the complete shared-string-optimized source
`67ab8dd54bcc17697df755b5b69bb36316056827` as its baseline. Its saved Excel
assembly SHA-256 is `E3358E58765B7060E8341AC361FB7FF09855A19E55E235AA5DC1E80483CB849C`.
Keep the four existing report inputs and all 2,048 output checks unchanged.
The candidate adds lookup modes, INDEX references/arrays and named expressions;
reference-backed lookup sources also avoid eager range materialization.

Compare that report before/after, then compare 256 linear and binary XLOOKUPs
over identical sorted numeric sources at 512 and 8,192 keys using
`FormulaLookupScaleBenchmarks`. Setup and complete output validation remain
outside timing. Run the complete matrix on each fixed processor domain. Accept
only correct outputs; investigate material common-case regressions. Stop after
the measured source-resolution and pass-catalog changes are qualified, without
adding speculative persistent indexes or expression caches.
