# Excel and CSV tabular data performance

## Workload and acceptance

I compare complete 50,000-row operations, with the same work for every engine.
Excel uses a unique name, integer ID, date and double; CSV uses the existing
mixed ten-column fixture. Reads return owned strings and consume typed values.
Object parsing constructs one record per row. Writers include row generation,
serialization, finalization and the complete in-memory byte array. Setup checks
every value through independent readers; measured reads check counts and checksums.
Invalid output has no usable timing. Third-party dependencies remain in the
opt-in comparison projects.

I target the largest avoidable costs first. I keep a change when correctness
holds and the affected comparison shows a material timing or allocation gain.
I stop after removing the demonstrated large costs and qualifying the stable
candidate; timing noise does not justify another API, dependency or speculative
fast path. These runs do not establish a portable ranking.

## Measurement cohort: 2 November 2026

The baseline is `3cc5e134b9`, after the omitted-reference fix and comparison
adapters. The starting integration revision is `3e805c1bcb`. I use BenchmarkDotNet
0.15.8 in process: eight warmups, twelve measurements, sixteen invocations,
unroll one, no outlier removal. Every build, test and benchmark uses the shared
heavy-run lock, with builds limited to two MSBuild workers.

The host is an Apple M4 with ten cores and 24 GiB RAM, macOS 27.0.1 (26A434),
Arm64 .NET 10.0.14, SDK 10.0.403, concurrent workstation GC. The ten cores
comprise four performance and six efficiency cores; AC low-power mode is off.
Process priority is normal: the attempted BenchmarkDotNet priority elevation is unavailable on this
host. No processor affinity is applied; the operating system schedules work
across the M4's heterogeneous cores. I retain variation and do not claim a quiet
host or equivalence from overlapping point estimates.

For reader comparisons I pin baseline-produced files and their SHA-256 hashes,
then use the same benchmark sources and options with baseline and candidate
engine DLLs. This avoids changing the input when the native writer improves. Writer and CSV
baselines use the original full matrix. The normalized evidence records the
package versions, source and assembly identities, input hashes, per-iteration
samples, medians, allocation and validation outcomes.

## Profile findings and changes

- The mapping fast path shared a method with a fallback lambda that captures the
  reader. The closure was allocated even when the fast mapper returned first.
  I reuse the existing slow helper, keeping the constructor/AOT fallback intact.
- The workbook reader entered its large lazy-open method on every accessor and
  repeated lifecycle/cancellation checks. I keep those checks in the accessor
  and call the lazy method only before the sheet has opened.
- Ordinary contiguous BIFF strings used a StringBuilder and a temporary decoded
  string before producing the owned result. I decode a fitting segment directly;
  strings crossing CONTINUE records retain the existing incremental decoder.
  BIFF compressed Unicode uses Latin-1, preserving byte-to-character semantics.
- Native XLS tabular writing rejected dates and expanded the entire table into
  editable Open XML cells. The baseline allocated about 582 MiB per export.
  I capture date serials with the workbook's date system, reuse the existing
  insertion date-style owner and carry the BIFF style through cached and
  uncached emission. Validation, mutable-source fingerprints, cancellation,
  Index/DBCell layout and fallback for unsupported workbook state remain intact.

CPU and allocation profiles are diagnostic runs, separate from the reported
measurement cohort. The separate XLSX profile did not establish a large cost in
the synchronous-to-asynchronous row adapter. I leave that path unchanged rather
than adding a second row loop for an unmeasured gain. Encoding, compression,
package buffering and row generation remain part of that complete export.

## Results

The [normalized evidence](excel-tabular-data-2026-11-02.json) contains all 36
before/after pairs, twelve samples per side, managed allocation, assembly hashes,
input hashes and exact third-party versions/licenses. It distinguishes the
fixed-input reader control from the initial writer and CSV baseline. Every lane
passes its preflight; no invalid package is ranked.

Native XLS date exports improve from 856.816 to 23.564 ms, about 36 times faster,
with allocation falling from 581.620 to 38.625 MiB. Materialized XLS reads fall
from 7.553 to 5.214 ms and 7.836 to 3.269 MiB; object parsing falls from 10.503 to
7.499 ms and 11.270 to 5.558 MiB. XLSX and XLSB object parsing each eliminate about
1.14 MiB per operation. Raw XLSX/XLSB read allocation was already close to the
matched owned-string comparison and stays essentially unchanged.

The XLSX writer is unchanged: 9.928 versus 9.884 ms, 9.066 MiB on both runs.
XLSB writing changes from 19.198 to 16.030 ms with identical 15.177 MiB allocation,
despite no change to that writer. That variation is a reason to avoid attributing
small timing differences to this patch. CSV allocation is unchanged; the complete
per-library results include those timing movements too. Initial and fixed-input
baseline reader launches are retained separately as evidence of launch variation.

## Follow-up cohort: 3 November 2026

The baseline is `9db3761b50`; the measured candidate is `e00a512a8a`, which also
merges the current performance base. Harness, options, pinned reader inputs and
host match the 2 November cohort. The [normalized evidence](excel-tabular-data-2026-11-03.json)
holds all 36 pairs with samples, allocation and assembly hashes. Other serialized
heavy runs shared the host between launches, so I treat movements under about 5%
in unchanged code as launch variation.

Profiles showed these hidden costs, and I removed them without changing the API,
the validation boundaries or the returned values:

- XLSB text decoded through `Encoding.Unicode`, whose character counting took
  about a seventh of a raw read. Text without surrogate code units now copies
  directly; any surrogate keeps the encoder's pairing and U+FFFD rules.
- XLSB discovery sent every inline string through the general cursor validator.
  A fitting `BrtCellSt` now passes the fused checks; other shapes keep the exact
  cursor errors.
- The XLS shared-string writer made several stream calls, limit checks and an
  encoder allocation per string; it took about a quarter of an XLS write. It now
  writes into one bounded buffer per record payload with the same segment limits.
  Cached rows no longer zero per-row stack buffers, and table entries are values.
- The XLSX string cache kept probing eight full slots for unique text. It now
  samples its hit rate and pauses lookups while fewer than one in sixteen repeat.
- The tabular readers reset small per-row arrays through the general
  `Array.Clear`; they now use typed span clears.

| Lane, median ms | Before | After | ExcelReader.NET |
|---|---:|---:|---:|
| XLSX read | 10.314 | 9.957 | 6.373 |
| XLSX typed parse | 11.987 | 11.278 | 8.083 |
| XLSB read | 8.058 | 6.101 | 4.764 |
| XLSB typed parse | 9.968 | 7.766 | 6.057 |
| XLS read | 5.387 | 5.244 | 4.627 |
| XLS typed parse | 7.569 | 7.527 | 7.519 |
| XLSX write | 9.924 | 9.924 | 7.974 |
| XLSB write | 16.380 | 16.214 | 5.448 |
| XLS write | 23.959 | 17.749 | 3.941 |

XLS write allocation falls from 38.622 to 36.194 MiB; every other OfficeIMO lane
keeps its allocation. Sylvan, ExcelDataReader, SpreadCheetah and all CSV lanes
stay within launch variation. In the XLBench shape comparison, edited-read moved
from 659.5 to 708.5 ms and calc from 43.6 to 45.8 ms; two interleaved
baseline/candidate reruns put both within 2% of each other. Every other category
moved by 4% or less.

## Typed mapping and compact numbers: 3 November 2026

The retained changes map typed rows through the active worksheet reader and write
only bit-exact RK numbers in XLS and XLSB. Binary readers retain complete validation
before row delivery, used-range widening and stale-dimension trimming. Excel typed
tables reject unsupported captions instead of silently discarding them.

The original baseline is `5d33057990`. The earlier mixed cohort does not isolate
these retained changes, so it cannot establish their independent throughput gain.
The direct mapper keeps per-cell compatibility checks and the outer-reader conversion
fallback. RK encoding is used only when decoding reproduces the original double bits.

## Compatible writer follow-up: 3 November 2026

Fresh synchronous direct-tabular XLS stream saves prepare the BIFF buffer and compound
layout before touching the destination, then use the existing compound serializer
to emit directly. Ordinary memory destinations reserve the exact serialized size,
as the previous single staged write did. Imported workbooks, asynchronous or
cancellable saves, and timing callbacks retain their existing staging boundary.
Source enumeration and selectors still run during `WriteData`; their values remain
buffered until saving.

The first 50,000-row measurement reduces XLS write allocation from 36.194 MiB at
`5d33057990` to 30.072 MiB. The initial direct-write attempt without the exact
reservation used 35.674 MiB: repeated destination growth consumed much of the saved
copy. That version is superseded by the exact reservation. Complete output is
validated by the strict reader and independently by ExcelDataReader.

The [normalized compatible cohort](excel-tabular-compatible-2026-11-03.json) records
all 25 ExcelReader.NET, Sylvan, ExcelDataReader and SpreadCheetah lanes in each of
two interleaved baseline/candidate pairs. The implementation is `40734fa5ac`.
The same baseline-produced files and SHA-256 hashes are used throughout. Each cell
below keeps the first and second launch separate; allocation is representative of
both pairs.

| Lane | Before ms, pair 1 / 2 | After ms, pair 1 / 2 | Before → after MiB |
|---|---:|---:|---:|
| XLSX read | 10.549 / 10.169 | 10.134 / 9.927 | 2.274 → 2.274 |
| XLSX typed parse | 12.701 / 11.465 | 11.021 / 10.858 | 4.562 → 4.562 |
| XLSB read | 6.434 / 6.323 | 6.380 / 6.743 | 2.284 → 2.285 |
| XLSB typed parse | 8.023 / 8.023 | 7.090 / 7.056 | 4.574 → 4.574 |
| XLS read | 6.091 / 5.353 | 5.829 / 5.533 | 3.268 → 3.269 |
| XLS typed parse | 8.245 / 7.427 | 6.769 / 6.639 | 5.558 → 5.558 |
| XLSX write | 10.220 / 9.912 | 10.086 / 9.933 | 9.065 → 9.066 |
| XLSB write | 16.458 / 16.064 | 16.683 / 16.687 | 15.173 → 15.470 |
| XLS write | 17.602 / 17.694 | 18.196 / 18.727 | 36.193 → 30.072 |

Typed parsing is 5–18% faster in both pairs. Raw reads have no consistent material
gain. XLS writes save 17% of managed allocation; their raw medians are 3–6% slower,
while the unchanged ExcelReader.NET XLS writer is 7–9% slower in those same pairs.
I retain the allocation reduction without claiming an XLS throughput gain. The
retained XLSB RK output adds about 0.30 MiB of writer allocation; the complete
writer still trails ExcelReader.NET.

All fifteen XLBench shapes complete their validation. Timing shifts in unchanged
paths remain visible in the evidence. Two targeted interleaved repetitions reverse
the direction of the calc difference: 103.0 → 112.3 ms, then 114.3 → 102.7 ms,
with 16.57 MiB in every run. Template allocation ranges from 7.88 to 7.90 MiB on
both sides. These runs do not establish a persistent timing regression or a quiet
host result.

The matrix uses the existing runner, with the pinned input root supplied through
`OFFICEIMO_EXCEL_GENERATED_INPUT`. Set `EVOTEC_SCRATCH_ROOT` to the local scratch directory:

```sh
lockf -k "${EVOTEC_SCRATCH_ROOT:?Set EVOTEC_SCRATCH_ROOT}/heavy-run.lock" \
  dotnet OfficeIMO.Excel.Benchmarks/bin/Release/net10.0/OfficeIMO.Excel.Benchmarks.dll \
  --filter '*ExcelReaderGenerated*' --inProcess --warmupCount 8 --iterationCount 12 \
  --invocationCount 16 --unrollFactor 1 --outliers DontRemove --artifacts <output>/bdn
lockf -k "${EVOTEC_SCRATCH_ROOT:?Set EVOTEC_SCRATCH_ROOT}/heavy-run.lock" \
  dotnet OfficeIMO.Excel.Benchmarks/bin/Release/net10.0/OfficeIMO.Excel.Benchmarks.dll \
  --compare-xlbench-shapes 9 2
```

Two read experiments are not retained:

- Inflating the same validated XLSB and XLSX parts into an existing buffer with
  81,920-byte, 256 KiB, 1 MiB and whole-part reads keeps allocation unchanged.
  Medians range from 2.127 to 2.142 ms for XLSB and 2.585 to 2.619 ms for XLSX.
  The chunk-size change does not earn a production change.
- A dedicated validated XLSB inline-text delivery branch avoids a general cell
  dispatch but shows no material read improvement or allocation reduction. The
  binary readers therefore remain byte-for-byte identical to the baseline source.

The existing RK boundary test checks decoded double bits across both binary formats
and resaves. An additional negative-zero control through `Cell.SetValue` returns
positive zero on the original baseline and the candidate, including XLSX. That
existing normalization is preserved rather than changing the cell-write contract.

## Shared-string storage follow-up: 3 November 2026

I profile the integrated implementation at `6c56a09f12` and compare it with
`5e87f1f784`, using the same benchmark sources, pinned reader inputs and settings
as the compatible cohort above. I retain two interleaved baseline/candidate
pairs for all 36 Excel and CSV lanes. Rebuilt assemblies stamp the Git revision
into their metadata; the evidence retains each actual assembly hash rather than
claiming byte-identical harness DLLs.

The read profiles still spend substantial time constructing owned strings and
mapping typed rows. A restarted XLSB row plan can recognize a stable suffix after
text lengths change, but its targeted runs do not show a repeatable material gain.
A part-wide XML normalization flag likewise does not improve XLSX reads. I remove
both experiments and their associated test extension. Complete binary discovery,
used-range handling, XML validation and the existing read implementations remain
unchanged.

The binary write profiles show source-row materialization, boxed selector values,
string generation, dictionary growth and shared-string storage. The direct XLS
builder keeps both a dictionary and a growing entry list for the same unique
strings. I use the dictionary's explicit numeric indexes to construct one exact
entry array instead. Newer targets also insert through one dictionary lookup;
netstandard2.0 and net472 retain ordinary dictionary operations. Entry order does
not depend on dictionary enumeration order. Duplicate strings still contribute
to the total count, while each distinct string has one entry.

The [normalized evidence](excel-tabular-direct-sst-2026-11-03.json) includes twelve
samples per launch, allocation, source and assembly identities, input hashes,
package versions and the rejected read experiments. The table keeps both launches
separate. Allocation uses the second pair; both pairs show the same XLS saving.

| Lane | Library | Before → after ms, pair 1 | Before → after ms, pair 2 | Before → after MiB |
|---|---|---:|---:|---:|
| XLSX read | ExcelDataReader | 62.486 → 63.749 | 62.185 → 62.624 | 33.881 → 33.881 |
| XLSX read | ExcelReader.NET | 6.322 → 6.370 | 6.667 → 6.406 | 2.217 → 2.217 |
| XLSX read | OfficeIMO | 10.145 → 10.147 | 10.146 → 10.120 | 2.274 → 2.274 |
| XLSX read | Sylvan | 29.660 → 29.140 | 29.418 → 29.130 | 2.546 → 2.546 |
| XLSX parse | ExcelReader.NET | 8.256 → 7.773 | 8.241 → 7.856 | 4.506 → 4.506 |
| XLSX parse | OfficeIMO | 10.963 → 10.977 | 10.966 → 11.065 | 4.562 → 4.562 |
| XLSX write | ExcelReader.NET | 7.946 → 8.096 | 8.019 → 8.028 | 9.008 → 9.008 |
| XLSX write | OfficeIMO | 9.857 → 9.902 | 9.939 → 9.867 | 9.066 → 9.066 |
| XLSX write | SpreadCheetah | 11.924 → 11.944 | 11.645 → 12.185 | 11.135 → 11.135 |
| XLSB read | ExcelDataReader | 23.578 → 23.053 | 23.660 → 22.601 | 18.702 → 18.702 |
| XLSB read | ExcelReader.NET | 4.765 → 4.708 | 4.878 → 4.797 | 2.217 → 2.217 |
| XLSB read | OfficeIMO | 6.539 → 6.359 | 6.246 → 6.106 | 2.285 → 2.285 |
| XLSB read | Sylvan | 5.943 → 5.716 | 5.812 → 5.680 | 2.464 → 2.464 |
| XLSB parse | ExcelReader.NET | 6.102 → 6.181 | 6.201 → 6.315 | 4.506 → 4.506 |
| XLSB parse | OfficeIMO | 6.806 → 7.186 | 6.824 → 6.852 | 4.574 → 4.574 |
| XLSB write | ExcelReader.NET | 5.486 → 5.436 | 5.404 → 5.438 | 8.829 → 8.829 |
| XLSB write | OfficeIMO | 15.972 → 16.142 | 16.125 → 16.986 | 15.469 → 15.470 |
| XLS read | ExcelDataReader | 43.766 → 43.492 | 48.549 → 42.723 | 50.960 → 50.960 |
| XLS read | ExcelReader.NET | 4.816 → 4.619 | 4.706 → 4.654 | 3.762 → 3.761 |
| XLS read | OfficeIMO | 5.215 → 5.464 | 5.680 → 5.224 | 3.269 → 3.268 |
| XLS read | Sylvan | 5.071 → 5.083 | 5.647 → 5.162 | 2.716 → 2.716 |
| XLS parse | ExcelReader.NET | 7.477 → 7.621 | 7.689 → 7.689 | 6.051 → 6.051 |
| XLS parse | OfficeIMO | 6.620 → 6.739 | 6.656 → 6.647 | 5.558 → 5.558 |
| XLS write | ExcelReader.NET | 4.241 → 4.258 | 4.098 → 4.218 | 17.074 → 17.074 |
| XLS write | OfficeIMO | 18.315 → 16.641 | 17.949 → 16.423 | 30.072 → 28.836 |
| CSV read | CsvHelper | 18.461 → 18.288 | 18.181 → 18.199 | 21.529 → 21.529 |
| CSV read | Dataplat.Dbatools.Csv | 11.258 → 11.478 | 11.203 → 11.061 | 21.514 → 21.514 |
| CSV read | ExcelReader.NET | 7.533 → 7.508 | 13.407 → 8.019 | 21.508 → 21.508 |
| CSV read | OfficeIMO | 7.989 → 8.083 | 8.097 → 8.052 | 21.511 → 21.511 |
| CSV read | Sep | 4.645 → 4.626 | 4.704 → 4.581 | 20.835 → 20.835 |
| CSV read | Sylvan | 8.825 → 8.908 | 11.895 → 8.836 | 21.542 → 21.542 |
| CSV parse | CsvHelper | 32.378 → 32.271 | 32.248 → 32.234 | 32.274 → 32.275 |
| CSV parse | ExcelReader.NET | 14.007 → 13.793 | 14.230 → 14.009 | 16.779 → 16.779 |
| CSV parse | OfficeIMO | 10.795 → 10.994 | 10.819 → 10.989 | 16.782 → 16.782 |
| CSV write | ExcelReader.NET | 6.642 → 6.691 | 6.482 → 6.509 | 15.961 → 15.960 |
| CSV write | OfficeIMO | 6.919 → 6.870 | 6.418 → 6.343 | 16.009 → 16.009 |

XLS writes save 1.236 MiB per operation, about 4.1%, and their medians improve
by 9.1% and 8.5% in the two pairs. The XLSB writer is unchanged; its medians move
from 15.972 to 16.142 ms and from 16.125 to 16.986 ms. I do not attribute those
movements, or the small read movements, to the shared-string change. The unchanged
ExcelReader.NET CSV reader moves from 13.407 to 8.019 ms in the second pair,
which demonstrates substantial launch variation even with serialized heavy work.

These are owned-string read comparisons: ExcelReader.NET also allocates about
2.217 MiB for the unique XLSX/XLSB names. The earlier few-kilobyte figures for
span-only scans do not describe this materialized workload. OfficeIMO's matched
raw reads allocate about 2.274 and 2.285 MiB respectively. Throughput parity is
still incomplete. Source materialization and selector invocation order, boxed
values, complete part validation, fixed reader schemas and binary compression
remain the principal constraints described below.

Excel builds for netstandard2.0, net8.0, net10.0 and net472 with no warnings or
errors. The net10.0 osx-arm64 NativeAOT publish and smoke pass. Final net10.0
correctness runs use `Category!=Performance`: Excel.Tests passes 7,751 with five
opt-in/platform skips, including all twelve omitted-reference cases; CSV.Tests
passes all 938. The hand-authored ZIP/XML fixture and grid loader are unchanged.
An independent read-only review finds no actionable issues in the retained writer
change. No additional test is needed for this storage change: existing binary
format and reader-boundary checks cover its values, indexes and duplicate strings.
The net472 result is a cross-build, without a Windows runtime run.

## XLSB source buffering and typed capture: 3 November 2026

I profile `3bd89bcafe` and retain two changes in `8b98201b4c` and `231694d85e`.
The XLSB selector profile spends its largest managed samples materializing source
rows and projecting boxed selector results. I replace growth copies of the private
source list with bounded chunks. Enumeration still completes and disposes before
selectors run. Projection reads each privately captured row once and releases it
after its final selector; caller-owned indexed sources keep their per-selector
reads. Values are still captured before `WriteData` returns.

The XLSB writer also passes typed worksheet-buffer cells through `GetValue`, which
boxes each integer, date, double and Boolean before the native plan unboxes it.
I capture those existing typed cells directly, including sparse coordinates and
explicit empty text. The native plan remains a separate validated snapshot before
any destination write. Its date system, styles, cancellation, fallback and pooled
buffer clearing remain intact. The worksheet grid and cell/range API are unchanged.

I keep typed and object-row capture in separate methods. Dispatch inside the large
object-row method shows a 5–7% slower selector lane in the earlier mixed runs;
that version is superseded. The [normalized evidence](excel-xlsb-buffering-2026-11-03.json)
retains those runs, the source-only experiment, exact source/assembly hashes and
the final qualification. Candidate assemblies were compiled before committing;
source-file hashes bind them to the retained implementation. Harness sources are
the same, and the actual rebuilt DLL hashes are recorded.

The final launch order is candidate, baseline, candidate, with twelve samples per
launch. This is one baseline bracketed by two candidates, not two independent
baseline pairs. Settings remain eight warmups, sixteen invocations, unroll one and
no outlier removal on the same M4 with ordinary OS scheduling. All eight cases
validate every header and value through the strict reader and independent
ExcelDataReader before timing. The buffered-cell diagnostic compares OfficeIMO
revisions separately; the engine comparison lanes retain their matched workload.
Every operation includes row generation, serialization, finalization and its
complete byte array. Each operation writes 50,000 rows with four typed columns.

| Write lane | Library | Baseline ms | Candidate 1 / 2 ms | Before → after MiB |
|---|---|---:|---:|---:|
| XLSB selectors | OfficeIMO | 16.215 | 16.069 / 16.120 | 15.469 → 14.875 |
| XLSB typed cells, diagnostic | OfficeIMO | 17.398 | 15.371 / 15.035 | 16.171 → 12.737 |
| XLSB | ExcelReader.NET 7.3.4 | 5.390 | 5.516 / 5.403 | 8.829 → 8.829 |
| XLS selectors | OfficeIMO | 16.282 | 15.165 / 15.053 | 28.834 → 28.241 |
| XLS | ExcelReader.NET 7.3.4 | 4.483 | 4.612 / 4.347 | 17.076 → 17.076 |
| XLSX streaming | OfficeIMO | 9.969 | 9.667 / 9.658 | 9.066 → 9.066 |
| XLSX | ExcelReader.NET 7.3.4 | 7.895 | 7.959 / 7.881 | 9.007 → 9.007 |
| XLSX | SpreadCheetah 1.28.0 | 11.207 | 11.479 / 11.438 | 11.135 → 11.135 |

Typed-cell XLSB exports allocate 3.43 MiB less, about 21%, and their medians are
11.7–13.6% faster. The selector route saves about 0.59 MiB, or 3.8%; its final
medians differ by less than 1%, so I claim an allocation reduction without a
material selector throughput gain. Initial selector timings range from a faster
source-only experiment to slower mixed runs. Those movements and the unchanged
peer lanes remain visible in the evidence; this is no portable ranking.

The public `Func<T, object?>` selectors still box value-type results. I reduce the
source-buffer growth and overlapping source-row lifetime, and remove the later
reboxing of already typed worksheet cells. I retain captured selector values and
the separate native snapshot rather than borrowing mutable inputs during writes.

The focused final run passes 318 cases. The new public-boundary test passes both
against the baseline engine and this implementation: enumeration/disposal precedes
row-major selectors across 4,097 rows, post-call mutations do not alter captured
values, and repeated XLSB saves never invoke selectors again. Existing indexed
source, scalar/date, sparse/blank, cancellation, failure and pool-clearing tests
remain intact. Final net10.0 correctness runs pass 7,752 Excel tests with five
opt-in/platform skips and all 938 CSV tests, including the twelve omitted-reference
cases. All four framework builds have zero warnings/errors; NativeAOT publish and
smoke pass. An independent read-only review finds no actionable issues. The net472
result is a cross-build without a Windows runtime run.

## Compressed XLSB snapshots: 3 November 2026

The baseline is `cc1dd3e42e`; `11d6fff204` replaces the native cell snapshot for
privately captured flat values exported to an ordinary empty `MemoryStream`.
I validate and serialize those values once into an owned compressed package,
then publish it after styles, shared strings and ZIP finalization succeed.
Fixed pooled chunks avoid growth copies, and the destination reserves the exact
completed package length. Every chunk is cleared and returned. Other destinations,
cancellable saves and timing callbacks retain their existing path. Selector timing,
the public API, worksheet grid and runtime dependencies are unchanged.

The profile still shows source materialization, boxed selectors, string generation
and memory-stream growth. Managed sampling does not isolate native DEFLATE CPU
time, so I make no compression-algorithm throughput claim. Compression remains
`Fastest`. For the selector workload, the worksheet stays 5,127,936 bytes before
compression and 841,866 bytes after compression; the whole package stays 845,732
bytes. Its worksheet SHA-256 is identical in the baseline and both candidate
launches. The allocation reduction does not trade compression ratio or data.

The first compressed-staging experiment uses a growing `MemoryStream` and raises
XLSB allocation from 14.875 to 16.515 MiB without a material timing improvement.
I discard that storage choice. The [normalized evidence](excel-xlsb-compressed-snapshot-2026-11-03.json)
retains that result, the pooled trial, profiles, actual assembly/source hashes,
package/worksheet sizes and all samples. The final runs use the exact same
benchmark binary with candidate, baseline, candidate ordering: one baseline
bracketed by two candidates. Each launch keeps twelve samples, eight warmups,
sixteen invocations and all outliers. Package generation, row generation,
serialization, finalization and the complete byte array remain inside the
50,000-row operation. Every lane passes strict and independent all-value checks.

| Write lane | Library | Baseline ms | Candidate 1 / 2 ms | Before → after MiB |
|---|---|---:|---:|---:|
| XLSB selectors | OfficeIMO | 16.957 | 16.184 / 15.280 | 14.874 → 13.368 |
| XLSB typed cells, diagnostic | OfficeIMO | 16.040 | 15.621 / 14.637 | 12.737 → 12.737 |
| XLSB | ExcelReader.NET 7.3.4 | 5.531 | 5.598 / 5.489 | 8.829 → 8.829 |
| XLS selectors | OfficeIMO | 16.111 | 14.466 / 14.314 | 28.241 → 28.242 |
| XLS | ExcelReader.NET 7.3.4 | 5.136 | 5.006 / 4.486 | 17.075 → 17.076 |
| XLSX streaming | OfficeIMO | 10.200 | 10.544 / 9.761 | 9.066 → 9.066 |
| XLSX | ExcelReader.NET 7.3.4 | 8.178 | 8.061 / 8.046 | 9.007 → 9.007 |
| XLSX | SpreadCheetah 1.28.0 | 11.510 | 11.634 / 11.367 | 11.135 → 11.135 |

The retained XLSB selector route saves about 1.51 MiB per operation, or 10.1%.
Its medians improve by 4.6–9.9%, but unchanged XLS code also moves by about 10–11%.
I therefore keep the allocation result without attributing all timing movement
to the snapshot change. The typed-cell route is unchanged in this pass. The
benchmark preflight records package size, worksheet size, compressed length and
worksheet hash; these checks run outside the timed operation.

The existing late-overlong-text test now also exercises a bounded flat snapshot
through public `WriteData` and `Save`. The fresh destination keeps zero length,
position and capacity after failure. All four cases pass against the baseline
engine too. Final validation passes 182 focused cases, 7,754 net10.0 Excel tests
with five opt-in/platform skips, and all 938 CSV tests. All twelve omitted-reference
cases pass. netstandard2.0, net8.0, net10.0 and net472 builds have zero warnings/errors;
osx-arm64 NativeAOT publish and smoke pass. Independent read-only review finds
no actionable issues. net472 remains a cross-build without a Windows runtime run.

## XLSB discovery after a row-layout change (3 November)

I stop feeding the optional XLSB row-layout builder once discovery proves that
it cannot produce a delivery plan. The builder's lifetime replaces its disabled
flag. The complete worksheet-validation pass keeps every framing, coordinate,
style, string, cancellation and resource-budget check; only later no-op layout
observations disappear. Stable layouts retain the same delivery plan. The change
is confined to two private reader files.

Sampling still points to string materialization, inflation and XLSX indexing as
larger costs. Duplicate date conversion is not a leading sampled cost, so I leave
it unchanged. Managed sampling does not isolate native inflater throughput. The
[normalized evidence](excel-xlsb-discovery-2026-11-03.json) retains the profile
summaries, source and assembly hashes, inputs and every measured sample.

The final bracket uses one baseline between two candidates, the identical Excel
benchmark binary, eight warmups, twelve samples, sixteen invocations and all
outliers. Excel uses the pinned four-column input; CSV uses its existing
mixed ten-column workload and the identical unchanged binary in all three
launches. Every participant validates all values in preflight and produces or
consumes the complete output during timing. The typed-cell XLSB write remains a separate
diagnostic outside the comparison ranking.

XLSB direct reading improves by 2.0–2.5% in this cohort; typed parsing improves by
1.0–1.8%. Allocation is effectively unchanged. The earlier 32-invocation trial
improves parsing by 2.7–3.3% but shows no direct-read improvement; that trial is
retained too. Unchanged controls also move, especially CSV writing. I make no
portable ranking or broader read/write improvement claim.

Excel controls and changed lanes (50,000 rows):

| Lane | Library | Before ms | Candidate 1 / 2 ms | Before → candidate 1 / 2 MiB |
|---|---|---:|---:|---:|
| XLSB typed cells (diagnostic) | OfficeIMO | 14.933 | 14.760 / 14.950 | 12.737 → 12.737 / 12.737 |
| XLSX parse | OfficeIMO | 10.762 | 10.797 / 10.986 | 4.563 → 4.563 / 4.563 |
| XLSX parse | ExcelReader.NET | 8.298 | 8.160 / 8.355 | 4.506 → 4.506 / 4.506 |
| XLS parse | OfficeIMO | 6.686 | 6.647 / 6.648 | 5.558 → 5.558 / 5.558 |
| XLS parse | ExcelReader.NET | 7.643 | 7.652 / 7.598 | 6.051 → 6.051 / 6.051 |
| XLSB parse | OfficeIMO | 7.080 | 7.006 / 6.951 | 4.574 → 4.574 / 4.574 |
| XLSB parse | ExcelReader.NET | 6.220 | 6.286 / 6.012 | 4.506 → 4.506 / 4.506 |
| XLSX read | OfficeIMO | 9.954 | 10.002 / 10.002 | 2.274 → 2.274 / 2.274 |
| XLSX read | Sylvan | 30.281 | 30.199 / 30.761 | 2.546 → 2.546 / 2.546 |
| XLSX read | ExcelDataReader | 67.216 | 65.489 / 65.704 | 33.881 → 33.881 / 33.881 |
| XLSX read | ExcelReader.NET | 6.414 | 6.390 / 6.387 | 2.217 → 2.217 / 2.217 |
| XLS read | OfficeIMO | 5.253 | 5.266 / 5.257 | 3.269 → 3.269 / 3.269 |
| XLS read | Sylvan | 5.154 | 5.148 / 5.103 | 2.717 → 2.716 / 2.716 |
| XLS read | ExcelDataReader | 42.652 | 42.842 / 42.769 | 50.962 → 50.962 / 50.962 |
| XLS read | ExcelReader.NET | 4.607 | 4.649 / 4.599 | 3.762 → 3.762 / 3.762 |
| XLSB read | OfficeIMO | 6.347 | 6.218 / 6.190 | 2.284 → 2.284 / 2.284 |
| XLSB read | Sylvan | 5.649 | 5.644 / 5.620 | 2.464 → 2.464 / 2.464 |
| XLSB read | ExcelDataReader | 22.671 | 22.726 / 22.776 | 18.702 → 18.702 / 18.702 |
| XLSB read | ExcelReader.NET | 4.776 | 4.730 / 4.777 | 2.217 → 2.217 / 2.217 |
| XLS write | OfficeIMO | 14.931 | 14.869 / 14.985 | 28.241 → 28.241 / 28.241 |
| XLS write | ExcelReader.NET | 4.233 | 4.411 / 4.217 | 17.075 → 17.075 / 17.075 |
| XLSB write | OfficeIMO | 15.125 | 15.232 / 15.260 | 13.368 → 13.368 / 13.369 |
| XLSB write | ExcelReader.NET | 5.471 | 5.629 / 5.492 | 8.829 → 8.829 / 8.829 |
| XLSX write | OfficeIMO | 9.898 | 9.914 / 9.929 | 9.066 → 9.066 / 9.066 |
| XLSX write | ExcelReader.NET | 8.011 | 7.975 / 8.029 | 9.008 → 9.008 / 9.008 |
| XLSX write | SpreadCheetah | 11.334 | 11.349 / 11.829 | 11.135 → 11.135 / 11.135 |

CSV controls and changed lanes (50,000 rows):

| Lane | Library | Before ms | Candidate 1 / 2 ms | Before → candidate 1 / 2 MiB |
|---|---|---:|---:|---:|
| CSV read | OfficeIMO | 8.039 | 8.043 / 8.011 | 21.511 → 21.511 / 21.511 |
| CSV read | ExcelReader.NET | 7.497 | 8.143 / 7.449 | 21.508 → 21.508 / 21.508 |
| CSV read | Sylvan | 8.890 | 8.850 / 8.862 | 21.542 → 21.542 / 21.542 |
| CSV read | Sep | 4.576 | 4.575 / 4.656 | 20.835 → 20.835 / 20.835 |
| CSV read | CsvHelper | 18.352 | 18.295 / 18.259 | 21.529 → 21.529 / 21.529 |
| CSV read | Dataplat | 10.980 | 10.941 / 10.970 | 21.514 → 21.514 / 21.514 |
| CSV parse | OfficeIMO | 10.793 | 10.837 / 10.772 | 16.782 → 16.782 / 16.782 |
| CSV parse | ExcelReader.NET | 13.857 | 13.826 / 13.758 | 16.779 → 16.779 / 16.779 |
| CSV parse | CsvHelper | 32.545 | 32.235 / 32.235 | 32.274 → 32.274 / 32.274 |
| CSV write | OfficeIMO | 6.202 | 6.446 / 6.933 | 16.009 → 16.009 / 16.009 |
| CSV write | ExcelReader.NET | 6.590 | 6.548 / 6.614 | 15.960 → 15.961 / 15.961 |

Existing boundary tests cover stable layouts, layout changes and trailing empty
rows, so I add no tests. The final focused run passes 77 cases, including all
twelve hand-authored omitted-reference cases. The full net10.0 Excel suite passes
7,754 tests with five skips; CSV passes all 938 tests, with `Category!=Performance`.
netstandard2.0, net8.0, net10.0 and net472 builds pass without warnings or errors.
The net10.0 osx-arm64 NativeAOT publish and smoke checks pass. net472 remains a
cross-build without a Windows runtime run.

The public API, dependencies, selector timing, writers, row mapping and worksheet
grid loader are unchanged in this pass. The existing omitted-reference fix and
its reader-boundary regression remain intact. Read validation/indexing and
inflation, owned strings, selector boxing and binary write staging remain the
larger unresolved costs.

## XLSX validation and cached-row indexing (3 November)

I reuse a complete, qualified implicit row layout through a smaller scan. Every
cached cell tag must match exactly; every payload and row closure is checked,
including each shared-string index. The first row and changed layouts use the
existing parser. A failed shortcut restores the byte position and permanently
returns to ordinary discovery. Used bounds commit only after a complete row;
partial metadata is replaced by the existing retry and initialization paths.
The complete UTF-8/XML proof still runs before delivery. No validation check,
public API, runtime dependency, writer, selector timing or grid-loader behaviour
changes.

The baseline sampled read profile puts 25.9% in indexing, 24.7% in inflation,
5.5% in canonical XML proof and 5.2% in dimensionless range discovery. With the
shortcut, indexing's sampled time falls from 9,933 to 8,055 ms across the same
profile workload, about 18.9%; its share falls to 22.2%. The common-row Tier1
routine uses 3,352 code bytes and a 112-byte stack frame, against 3,796 and 176
for the former common path. The original core remains as fallback, so this is
not a whole-assembly size reduction. Managed sampling does not isolate native
inflater throughput; profiling timings are not comparison results.

The stable comparison uses the identical harness and fixed 50,000-row,
four-column XLSX input, eight warmups, twelve samples, thirty-two invocations
and all outliers. One baseline is bracketed by two candidates. All implementations
validate every value in preflight and consume complete rows/checksums in timing.
The [normalized evidence](excel-xlsx-indexing-2026-11-03.json) retains source and
assembly hashes, inputs, all samples, profiles, review and validation receipts.

| Lane | Library | Before ms | Candidate 1 / 2 ms | Before → candidate 1 / 2 MiB |
|---|---|---:|---:|---:|
| Typed parse | OfficeIMO | 10.693 | 10.460 / 10.396 | 4.562 → 4.562 / 4.562 |
| Typed parse | ExcelReader.NET | 8.141 | 8.354 / 8.163 | 4.506 → 4.506 / 4.506 |
| Cell-by-cell read | OfficeIMO | 9.944 | 9.719 / 9.805 | 2.274 → 2.273 / 2.274 |
| Cell-by-cell read | Sylvan | 28.760 | 29.377 / 28.768 | 2.546 → 2.546 / 2.546 |
| Cell-by-cell read | ExcelDataReader | 60.777 | 61.904 / 62.977 | 33.880 → 33.880 / 33.880 |
| Cell-by-cell read | ExcelReader.NET | 6.341 | 6.324 / 6.315 | 2.217 → 2.217 / 2.217 |

Direct reading improves by 1.4–2.3% and typed parsing by 2.2–2.8% in the final
cohort; allocation is effectively unchanged. The earlier narrow layout trial
has mixed typed-parse results and is retained. Other trials are discarded:
a provisional local cursor, exact-word tag matching, moved metadata writes and
a cold-parser JIT boundary do not give a consistent read gain. Their losses and
controls remain in the evidence. These host-dependent results do not establish
parity or a portable ranking.

The existing validation/rollback proof passes 479 focused cases, including all
twelve hand-authored omitted-reference cases. No test changes remain. Independent
read-only review finds no actionable issues. Final net10.0 correctness suites
with `Category!=Performance` pass 7,754 Excel tests with five skips and all 938
CSV tests. All four framework builds have zero warnings/errors; osx-arm64
NativeAOT publish and smoke pass. net472 remains a cross-build without Windows
runtime validation. Inflation, range discovery, owned strings and the retained
validation/indexing pass still account for the remaining XLSX read gap.

## Inflation profile and XLBench comparison (3 November)

I measured `28881d6e68` again before changing inflation or range discovery. The
[baseline evidence](excel-xlsx-inflation-baseline-2026-11-03.json) retains all six
lanes, twelve samples per lane, assembly/input identities and the sampled profile.
The pinned input, complete-value preflight and benchmark settings match the
preceding indexing cohort. These are baseline results; no new runtime change is
measured here.

| Lane | Library | Median ms | Allocated MiB |
|---|---|---:|---:|
| Cell-by-cell read | OfficeIMO | 9.966 | 2.274 |
| Cell-by-cell read | ExcelReader.NET | 6.782 | 2.217 |
| Cell-by-cell read | Sylvan | 28.954 | 2.546 |
| Cell-by-cell read | ExcelDataReader | 62.882 | 33.880 |
| Typed parse | OfficeIMO | 10.514 | 4.562 |
| Typed parse | ExcelReader.NET | 8.105 | 4.506 |

Inflation accounts for 26.3% of sampled read time and 22.8% of typed-parse time;
dimensionless range discovery accounts for 5.2% and 4.9%. Complete canonical XML
proof takes another 5.7% and 5.1%. Inflation already writes into an exact-size
pooled buffer and validates the declared decompressed length. A previous read-size
experiment gave no useful improvement. I retain complete validation; removing it
would change the accepted-input contract.

I also checked the [XLBench results](https://jafin.github.io/XLBench/results.html)
and source at `c9b94c5bf41fdb0013a66f4df32a84fd6eb79365`. The results use OfficeIMO
3.5.3 on Windows and a Ryzen 9 5950X, rather than this local source on an M4.
The earlier local [shape comparison](excel-tabular-compatible-2026-11-03.json)
already shows a substantially different picture:

| Shape | Local OfficeIMO ms / MiB | Local SDK ms / MiB |
|---|---:|---:|
| Write | 30.4 / 9.29 | 97.4 / 134.88 |
| Read all | 232.1 / 50.55 | 723.6 / 627.85 |
| 20 sheets | 62.2 / 13.08 | 300.6 / 184.22 |
| Report | 5.1 / 4.08 | 4.9 / 5.25 |
| Delete rows and recalculate | 8.1 / 3.87 | 17.8 / 9.99 |
| Insert columns and recalculate | 9.1 / 3.80 | 22.8 / 13.45 |
| Calculation | 46.0 / 16.57 | n/a |
| Template | 9.3 / 7.90 | 11.6 / 13.83 |
| Properties | 0.3 / 0.70 | 0.2 / 0.41 |

These are retained measurements from `40734fa5ac`, not a fresh run of the current
source or a ranking against XLBench's whole field. Saved outputs receive scenario
validation; the earlier read-all and 20-sheet lanes check consumption/counts but
still need complete-value preflight before supporting a validated ranking.
The stock inputs are deterministic stand-ins. The two hosts' timings cannot prove
that OfficeIMO wins. XLBench's OfficeIMO read opens an editable workbook, whereas
Sylvan streams rows. Its Sylvan and MiniExcel writers use a precomputed total;
formula-capable writers emit a SUM formula. IronXL's published row is an older
snapshot. I keep those differences explicit in any comparison.

There is no CSV timing scenario in the inspected XLBench results. The existing
matched CSV cohort above is the useful baseline: raw reading remains about
8.0 ms against Sep's 4.6 ms, while typed parsing is about 10.8 ms against
ExcelReader.NET's 13.9 ms and CsvHelper's 32.5 ms. The next comparison should
qualify full-value XLBench-shaped reads across the leading libraries, keep
editable-workbook and tabular contracts visible, and broaden equivalent CSV
writer coverage before claiming overall parity.

## Batched implicit-row discovery (3 November)

I count possible implicit row tags in batches on net8.0 and net10.0, rather than
restarting the span search for each row. The scan retains every middle-byte and
token-boundary check, the row-count cap and a scalar tail. Cancellation is checked
initially and every 64 KiB. Older targets and hardware without acceleration keep
the scalar implementation. The proposed range still requires complete indexing,
coordinate qualification and XML validation before delivery. The public API,
dependencies, grid loader, writers and selector timing are unchanged.

The implementation is `9a2876d8eb`, against `28881d6e68`. The
[normalized comparison](excel-xlsx-range-discovery-2026-11-03.json) retains the
initial baseline, all narrow trials and three stable six-lane launches. The
stable baseline uses the same harness, unchanged dependency DLLs, pinned input,
eight warmups, twelve measurements and thirty-two invocations as both candidates.
Every heavy command takes the shared lock separately, allowing other validation
to run between launches. No outliers are removed.

| Lane | Library | Before ms | Candidate 1 / 2 ms | Before → candidate 1 / 2 MiB |
|---|---|---:|---:|---:|
| Cell-by-cell read | OfficeIMO | 9.669 | 9.567 / 9.462 | 2.273 → 2.273 / 2.273 |
| Cell-by-cell read | ExcelReader.NET | 6.411 | 6.381 / 6.384 | 2.217 → 2.216 / 2.217 |
| Cell-by-cell read | Sylvan | 29.116 | 28.876 / 28.134 | 2.546 → 2.546 / 2.546 |
| Cell-by-cell read | ExcelDataReader | 61.512 | 61.053 / 59.839 | 33.880 → 33.880 / 33.880 |
| Typed parse | OfficeIMO | 10.454 | 10.342 / 10.295 | 4.562 → 4.562 / 4.562 |
| Typed parse | ExcelReader.NET | 8.052 | 7.865 / 8.152 | 4.506 → 4.506 / 4.506 |

Reading improves by 1.1–2.1% and typed parsing by 1.1–1.5% in the stable cohort;
allocation is effectively unchanged. Unchanged controls move too: Sylvan improves
by up to 3.4% and ExcelDataReader by up to 2.7%. I do not attribute every timing
movement to the change or claim parity with ExcelReader.NET. The earlier narrow
trial's slower typed-parse launch, 12.520 ms against 10.826 ms, remains in the
evidence alongside its second candidate at 10.296 ms.

The sampled raw-read share in the general byte-sequence span search falls from
5.57% to 1.67%. Attribution shifts into the range-creation caller after JIT
optimization, so I do not report a precise standalone after percentage for range
discovery. Missing inference frames do not mean zero cost. Native inflation stays
at about 26.6% and complete canonical XML proof at about 5.5%. I retain the existing
exact-size inflation path and all its length checks; the earlier read-size trials
provided no useful improvement.

All 479 focused reader cases pass. Independent read-only review finds no actionable
issues. No tests change. The final net10.0 suites with `Category!=Performance` pass
7,754 Excel tests with five skips and all 938 CSV tests, including all twelve
independently authored omitted-reference cases. All four framework builds have
zero warnings/errors; osx-arm64 NativeAOT publish and smoke pass. Test compilation
retains warnings in unchanged test sources. net472 remains a cross-build without
Windows or PowerShell 5.1 runtime qualification.

## Remaining costs and limits

Parity is incomplete; the earlier cohorts still show gaps against ExcelReader.NET
in XLSX/XLSB reads and typed parsing. The remaining read cost is structural rather
than one hidden hot spot. OfficeIMO proves the whole part before delivering the first row and gives the data reader a fixed field
count. It therefore makes a validation and indexing pass and then a delivery
pass; ExcelReader.NET reads each row once. Inflation remains a substantial part of the
XLSX and XLSB reads. Typed parsing qualifies every mapped cell on the worksheet
reader before the compiled mapper reads it.

The XLS and XLSB writers are the `WriteData` plus `Save` route. That route finishes source enumeration before
selectors run and retains their boxed values until saving. Privately enumerated
source rows are released as projection consumes them; caller-owned indexed sources
keep their original lifetime. Earlier stage profiling puts registration at about
7 ms of each write against 0.7 ms to generate the same rows. XLSB
deflate takes about another 5 ms; exact RK encoding reduces the numeric payload
where possible. Imported and asynchronous XLS saves retain compound-file staging.
Removing the buffering would
change when selectors and source enumeration run, so I leave that contract
unchanged here.

These are single M4 cohorts with OS scheduling across both core types. I do not
infer Windows, .NET Framework runtime, peak-memory or large-workbook results from
them.

## Correctness and compatibility proof

The independent omitted-reference fixture is authored as ZIP/XML, without a
third-party writer. Its single reader-boundary theory covers UTF-8/UTF-16,
streaming and snapshot mapping, loaded DOM, parallel rows and forward/backward
cell lookup. Before the fix, both grid-lookup cases fail at A2; the other eight
initial routes pass. The final twelve-case theory also checks read-only lookup
and DOM preservation. The fix only changes private coordinate lookup; the grid
loader and public cell/range/formatting API are unchanged. The focused omitted
and implicit-coordinate run passes 136 cases.

The native-date proof extends the existing insertion regression across XLS and
XLSB, both date systems and both insertion routes, and adds a date to the existing
uncached sparse-table case. The focused final run passes 32 cases. Every measured
writer also passes independent ExcelDataReader and strict-reader validation.

Excel builds for netstandard2.0, net8.0, net10.0 and net472; both comparison projects
also build for net8.0. The net10.0 osx-arm64 NativeAOT publish and smoke pass,
including constructor-bound record mapping. Final net10.0 correctness runs use
`Category!=Performance`: Excel.Tests passes 7,742 with five opt-in/platform skips;
CSV.Tests passes all 938. No Office or GUI process is used. Review is a structured
self-review; an independent reviewer was unavailable.

For the 3 November changes, the existing reader-boundary proof already covers the
changed behaviour, so I add no tests. XLSB inline and shared strings with paired
and unpaired surrogates and NUL keep their decoded values. Inline-string count
and truncation failures keep their discovery errors. Native XLS long ASCII and
Unicode strings that span CONTINUE records, cached formula strings and rich-text
shared strings keep their output, and the pooled-cache clearing contract still
holds. The focused XLS, XLSB, reader, shared-string and mapping run passes 2,349
cases with two skips. Excel builds for netstandard2.0, net8.0, net10.0 and net472.
