# Migration runtime comparison — 2026-11-07 UTC

All 1,848 measured samples pass complete validation across two Windows processor
domains. The largest Word 6 inspection and DBX eager-search cases take about half
the baseline median time on both domains. The five profiles allocate less on
their targeted paths; elapsed results vary by operation, scale and domain.

The [Domain A time](migration-runtime-2026-11-07/domain-a-comparison-time.md) and
[allocation](migration-runtime-2026-11-07/domain-a-comparison-allocation.md) tables,
and the [Domain B time](migration-runtime-2026-11-07/domain-b-comparison-time.md)
and [allocation](migration-runtime-2026-11-07/domain-b-comparison-allocation.md)
tables retain all 33 supported phases. Ratios are candidate/baseline; lower means
less time or allocation. Allocation values are means, and time ratios use medians.

## Observations and variation

For the largest cases, Domain A / Domain B observations are:

- Word 6 load/inspection median ratios are **0.489 / 0.505**, with allocation
  ratios **0.852 / 0.863**. Conversion/save allocation is **0.980 / 0.981**.
- DBX eager-search median ratios are **0.487 / 0.500**, with allocation
  **0.753 / 0.753**. Ordinary DBX materialization and EML serialization allocate
  approximately the same amount as the baseline.
- BIFF4 load and conversion/save allocation is **0.955–0.957** of baseline.
  Pages 2009 loading allocates **0.578 / 0.578**; conversion/save allocates
  **0.848 / 0.851**. WordML loading allocates **0.948 / 0.949** and
  conversion/save **0.939 / 0.939**.

Slower results remain visible. The 1,000-row BIFF load median is **1.170×** on
Domain A and **0.983×** on Domain B. On Domain B, large DBX save is **1.050×**,
large BIFF save **1.008×**, and large Pages save **1.013×**. Tail latency also
varies: small BIFF save P95 is **2.086×** on Domain B, and large Word save P95 is
**1.418×** on Domain A. These data do not establish an elapsed-time improvement
for every path.

Per-launch variation matters. On Domain A, the two large Pages save candidate
medians are 269.2 ms and 163.6 ms; the baseline medians are 182.9 ms and 202.6 ms.
Recorded shared-host CPU observations span **17–99%**. The
[launch manifest](migration-runtime-2026-11-07/launches.json), per-domain launch
summaries and unchanged raw CSVs retain this variation and every measured sample.

An initial WordML cache key serialized property elements and increased allocation.
The [initial observation](migration-runtime-2026-11-07/initial-wordml-observation.json)
records that rejected shape and the revision using already-created mapped XML.
Those five-sample selection observations are separate from the final ABBA pools.

## Method and validation boundary

The [workload contract](../../OfficeIMO.Migration.Benchmarks/README.md) defines
five synthetic shapes at small, medium and large scales: ASCII Word 6 shared-FKP
character runs, BIFF4W finding-heavy rows, Pages 2009 drawables sharing one raster,
indexed DBX messages, and WordML 2003 paragraphs with repeated property groups.
Input generation and complete ordered semantic, exact loss-diagnostic, saved
content and applicable Open XML schema validation occur outside timing. Baseline
and candidate have identical input and loss ledgers across all 33 phases;
independent-producer feature qualification is separate evidence.

The baseline is `b609d9788cd69c9af1861a0c4b2aad7181a03632`; the candidate is
`78b6b825fc5d90bb1aa5bb14767bb70a36abd709`. The
[bindings](migration-runtime-2026-11-07/bindings.json) record source-file hashes,
commit verification, product DLL/PDB hashes and runner identity. Candidate
binaries were built before that source commit; all 29 bound modified production
files were verified byte for byte against it. Both revisions use the exact same
workload DLL, SHA-256
`61A6B293FBD09F4D65DD66FD4857034A1677252DD0E5A2E4602B6450D8686690`.

Builds use SDK 10.0.403. Measurement executes inside PowerShell 7.6.9 on
.NET 10.0.14 with PSPublishModule 3.0.189; the build SDK does not select that
runtime. Each affinity cohort uses fresh processes in Baseline, Candidate,
Candidate, Baseline order. Each launch performs three warmups and seven measured
iterations, with rotated case ordering, collection before each iteration, Normal
priority and no outlier removal. Two launches provide 14 measured samples per
revision, phase and domain. PowerForge's summary, comparison and Markdown
renderer produce the tables; individual launch summaries accompany each pool.
There are 96 expected skipped records for non-DBX content-search combinations.

The Windows 11 host has an AMD Ryzen 9 9950X3D2, 16 cores and 32 logical
processors, using the High performance power plan. Native topology discovery
identifies two 96 MiB L3 domains: `0xFFFF` and `0xFFFF0000`. Host-load observations
describe normal shared activity; they neither filter samples nor establish
exclusive use. The validation and consumer build processes finished before this
matrix. Domains remain separate in the results.

## Reproduction and limits

Build the opt-in host once and copy that identical workload binary into separate
baseline and candidate runtime roots matching the recorded product/dependency
identities. Record rebuilt identities separately. In a fresh PowerShell 7.6+
process for every launch, run:

```powershell
./Build/Benchmarks/Run-MigrationBenchmarks.ps1 `
    -BinaryRoot $binaryRoot -ProductSource $sourceIdentity -OutputRoot $outputRoot `
    -Operation LoadInspect,ConvertSave,SearchContent `
    -WarmupCount 3 -IterationCount 7 -ProcessorAffinityMask $mask
```

Use all five profiles and three scales, a distinct output root for each launch,
and the recorded ABBA order for each mask. Retain unchanged sample CSVs and their
hashes. Use PowerForge's `BenchmarkSummaryService.Summarize/Compare` and
`BenchmarkMarkdownRenderer` for aggregation and table rendering.

Managed allocation volume includes the invocation boundary and host effects.
It does not measure retained, peak or native memory. The raw CSV's
`WorkingSetDeltaBytes` is an end-minus-start observation. These synthetic shapes
do not qualify every producer or historical file variant. Pages 2009 retains
its explicit partial editable reconstruction policy. Interpret each operation
and domain separately, including slower observations and launch variation;
this comparison establishes neither a universal speedup nor a portable
absolute performance budget.
