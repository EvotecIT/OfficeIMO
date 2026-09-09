# Excel and CSV text performance — 2026-09-07

OfficeIMO's CSV and direct XLSX writers spend less time escaping long text after
replacing repeated character writes with bulk copies of ordinary text. Dense
quotes and XML markup retain a scalar path. Public APIs, quoting rules, XML
sanitization, encoding, and package defaults are unchanged.

The improvement is strongest for long notes. Short exports and dense text have
less stable timing on this busy workstation. These measurements establish a
specific Windows baseline, not an overall library ranking.

## Before and after

The table reports median milliseconds per complete export. Each text fixture has
1,000 two-column rows. Values are normalized from batches of 32 exports, with
24 measured iterations after eight warmups. Before/after execution order rotates
within each iteration; no outliers are removed. Both binaries use the same
benchmark harness, inputs, and output checks in separate assembly load contexts.

| Workload | L3 domain 0: before → after | L3 domain 1: before → after | Observed reduction |
| --- | ---: | ---: | ---: |
| CSV long notes, `AsNeeded` | 1.81 → 0.69 ms | 1.47 → 0.58 ms | 61–62% |
| CSV long notes, `Always` | 3.16 → 0.95 ms | 2.66 → 0.68 ms | 70–74% |
| XLSX long plain text | 3.22 → 1.35 ms | 2.76 → 1.48 ms | 47–58% |
| XLSX long escaped text | 3.56 → 2.39 ms | 3.63 → 2.60 ms | 29–33% |

The generated notes contain long repeated character runs, Unicode, and embedded
punctuation. They exercise long-text serialization but compress more readily
than arbitrary application documents. The `4096` parameter describes the base
payload length; markers and row identifiers add characters.

The complete [before/after data](excel-csv-text-2026-09-07/before-after.json)
includes every measured batch from the final matrices and density confirmations.
It also records results that did not improve:

- The 25,000-row CSV control was effectively unchanged: median changes of 0%
  and +1.4% across the two domains.
- The 25,000-row XLSX control had 4.4–6.9% higher medians. Paired mean-difference
  intervals included zero on both domains, and earlier protocols changed the
  direction of the difference. A small regression is not ruled out.
- Dense CSV JSON and all-quotes fields improved under `AsNeeded` in most runs.
  `Always` was more sensitive to run conditions: the final matrix's all-quotes
  case was 23.6% slower on domain 1, while two subsequent confirmations were
  about 5.6% faster. JSON `Always` ranged from a small improvement to a 6.7%
  slowdown in the confirmations. Those results do not support a uniform gain.
- Short-note timings varied by domain and run. They do not establish a general
  short-row speedup or a strict no-regression guarantee.

The first bulk-copy candidate regressed on dense quotes and markup. The final
implementation switches to scalar escaping when special characters cluster;
the CSV StringBuilder loop is kept in a compact helper. The evidence above is
from that final implementation, not the rejected candidates.

## Comparison with other libraries

Native BenchmarkDotNet reports retain means, medians, dispersion, managed
allocations, and all selected cases:

- [CSV text export: OfficeIMO and CsvHelper](excel-csv-text-2026-09-07/csv-text.md)
- [XLSX text export: OfficeIMO and SpreadCheetah](excel-csv-text-2026-09-07/xlsx-text.md)
- [CSV all-field read: OfficeIMO, Sep, Sylvan, and ExcelReader.NET](excel-csv-text-2026-09-07/csv-read.md)
- [XLSX typed read: OfficeIMO, Sylvan, and ExcelReader.NET](excel-csv-text-2026-09-07/xlsx-read.md)

For long CSV notes under `AsNeeded`, OfficeIMO's native mean was 0.76/0.57 ms
across the two domains, versus CsvHelper's 2.04/1.08 ms. Managed allocation was
about 8.0 MiB versus 23.7 MiB per export. Short JSON fields favored CsvHelper in
this run. The larger long-text difference should not obscure that smaller case.

For long plain XLSX text, OfficeIMO's mean was 3.42/2.88 ms, versus
SpreadCheetah's 2.57/2.10 ms. OfficeIMO allocated about 245 KiB, versus 225 KiB.
Escaped text and dense markup showed wider dispersion and domain-dependent
ordering. SpreadCheetah remains a useful target for complete streaming export;
its documented scope is forward-only workbook creation.
[SpreadCheetah documentation](https://github.com/sveinungf/spreadcheetah)

OfficeIMO's long XLSX packages contained 59,419 bytes for plain text, 60,824 for
escaped text, and 125,539 for markup. The corresponding SpreadCheetah packages
contained 63,367, 64,789, and 132,384 bytes in this run. Each package reopened
successfully with every expected cell. OfficeIMO's before/after output lengths
were equal; this check does not assert byte identity of ZIP metadata.

The authentic 65,535-row, fourteen-column sales fixture gave OfficeIMO CSV means
of 16.22/15.72 ms, Sep 17.84/14.81 ms, ExcelReader.NET 22.97/21.41 ms, and Sylvan
19.19/17.61 ms. OfficeIMO and Sep each allocated about 34.2 MiB. Every field is
decoded to a string and included in the observation; a rows-only or span-only
scan measures a different result contract. Sep's published parser-throughput
experiments explicitly distinguish such workloads.
[Sep benchmark methodology](https://nietras.com/2025/06/17/sep-0-11-0/)

For the matching XLSX typed scan, OfficeIMO means were 74.82/77.85 ms,
ExcelReader.NET 83.97/101.23 ms, and Sylvan 158.30/191.57 ms. Managed allocations
were approximately 179, 31, and 649 KiB respectively. ExcelReader.NET exposes a
remaining allocation gap even where OfficeIMO's measured elapsed time was lower.
Broad confidence intervals prevent a universal ranking from this workstation.
[ExcelReader.NET documentation](https://github.com/GabrielMarquezMatte/ExcelReader)

Reader measurements were collected before the final CSV writer adjustment;
reader source was unchanged throughout this work. Benchmark-only dependency
updates include ExcelReader.NET/Arrow 3.0.1, LargeXlsx 2.0.2, and EPPlus 8.7.0.
CsvHelper 33.1.0, Sep 0.17.0, Sylvan.Data.Csv 1.4.4,
Sylvan.Data.Excel 0.5.8, and SpreadCheetah 1.28.0 were the measured versions in
the linked reports. Third-party libraries remain benchmark dependencies.

## Contracts and validation

CSV timing includes construction and full export to a `StringWriter`, not disk
I/O or UTF-8 file encoding. Setup reads every field and requires exact text
equality between writers, including headers, quote mode, Unicode, and newlines.

XLSX timing includes reader construction, XML serialization, ZIP compression,
and package finalization. It uses the existing compact tabular export options
`IncludeCellReferences = false` and `UseSharedStrings = false`; this change does
not change their defaults. Independent ExcelDataReader validation checks headers,
every cell, and the absence of extra rows. Shared-string and cell-reference
variants are covered by correctness tests rather than the timing table.

Regression tests cover empty and long values, quote density, custom and
multi-character delimiters, all quote modes, formula-injection escaping,
Unicode, valid whitespace, invalid XML control characters, and scan boundaries.
Final validation passed:

- CSV: 494 tests on .NET 10, 494 on .NET 8, and 366 on .NET Framework 4.7.2.
- Excel: 3,868 tests on .NET 10 with five opt-in skips, excluding two independently
  reproduced pre-existing image-baseline failures; focused .NET 8 and Framework
  suites passed 594 and 587 tests with one opt-in skip each.
- The final CSV adjustment also passed 16 Excel CSV-adapter tests. New escaping
  tests passed with hardware intrinsics disabled, and both product libraries
  compiled for .NET Standard 2.0. Both benchmark projects compiled for .NET 8.
- Exported public type/member comparison found no API changes. An independent
  read-only serialization review and targeted confirmation found no actionable
  issues.

The two image failures were `CroppedImageExportMatchesApprovedBaselines` and
`TransformedImageExportMatchesApprovedBaselines`. Repeating them with the
unmodified baseline Excel assembly produced the same 49,441/428,480 and
13,097/802,370 differing-pixel counts. Their baselines were not updated.

## Environment and reproduction

The workstation was in active use. Windows reported a Ryzen 9 9950X3D2 with
16 cores, 32 logical processors, and two 96 MiB L3 domains. Each run used one
domain: `0xFFFF` or `0xFFFF0000`, Normal process priority, and retained outliers.
Task-owned benchmarks did not run concurrently. These masks describe this
machine and must be rediscovered on another machine.

The runtime was .NET 10.0.11, SDK 10.0.111, BenchmarkDotNet 0.15.8, and Windows
11 build 26200.9168. Rotated confirmations used PowerForge through the installed
local PSPublishModule build in PowerShell 7.6.5. Native text jobs used six warmups,
twelve measured iterations, eight invocations per iteration, unroll factor one,
and one launch. Reader-job parameters appear in their native report headers.

Use the [CSV benchmark instructions](../../OfficeIMO.CSV.Benchmarks/README.md#text-export-and-quote-density)
and [Excel benchmark instructions](../../OfficeIMO.Excel.Benchmarks/README.md#text-export).
To reproduce this native text sample, use `--invocationCount 8 --unrollFactor 1
--warmupCount 6 --iterationCount 12 --launchCount 1 --outliers DontRemove` and
`--priority Normal`; on this machine add `--affinityMasks 0xFFFF,0xFFFF0000`.
Choose an explicit `--artifacts` folder. For reader runs select the methods in
the linked native reports and use their recorded invocation counts.

The [provenance manifest](excel-csv-text-2026-09-07/provenance.json) records source
commits, measured binary and fixture hashes, dependency versions, and environment.
Before/after compares baseline `d3aa130ac5` with CSV `b387d70e56` and Excel
`d4335a4686`. Rebuilding the baseline product with the candidate benchmark harness
keeps fixtures and validation identical. Keep each product's dependency set
identical when loading the two versions, and run the public benchmark entrypoints
with PowerForge's rotated order. Timing summaries divide batch durations by 32;
the JSON retains original batch samples so that normalization is inspectable.

No Linux/macOS timings, managed peak-memory measurements, or quiet-machine
performance budgets were established. Open work belongs to the
[product roadmap](../ROADMAP.md#document-format-depth): profile the existing
pooled UTF-8 writer, encoding copies, compression, and package finalization;
measure CSV file streams and typed projections; and settle small-control timing
with stable workloads across supported operating systems.
