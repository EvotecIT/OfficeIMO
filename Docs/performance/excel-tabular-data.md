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

## Single-pass binary readers: 3 November 2026

The XLSB and XLS readers now validate each record when `Read` reaches it instead of
proving the whole worksheet first. The XLSB and XLS writers store exact numbers as RK
values, and typed mapping reads the worksheet reader directly instead of through the
workbook wrapper. The baseline is `5d33057990`; harness, options and pinned reader inputs
match the earlier cohorts. Other agents' builds loaded this host heavily during both
interleaved before/after pairs (load averages from 7 to 134), and unchanged third-party
lanes moved by up to 40%, so I report allocation and timing directions only.

| Lane | Before MiB | After MiB | Timing in both pairs |
|---|---:|---:|---|
| XLSX, XLSB, XLS read | 2.274, 2.284, 3.269 | unchanged | within launch variation |
| XLSB typed parse | 4.574 | 4.574 | about 8% faster |
| XLSX typed parse | 4.562 | 4.562 | 5% to 14% faster |
| XLS write | 36.193 | 34.459 | within launch variation |
| XLSB write | 15.173 | 15.470 | within launch variation |

Sylvan, ExcelDataReader, SpreadCheetah and every XLBench shape keep their allocation.
The single pass did not make XLSB reads measurably faster. The old discovery pass was
already one fused scan, and its row plan let delivery skip framing; one validating pass
costs about the same. A profile of the new reader puts inflation at about a third of an
XLSB read (2.3 ms alone for the 5.7 MB part) and the validating row loop at about a
quarter. The remaining XLSB gap is per-record work, not a second pass. Typed parsing
spent about a tenth of its time qualifying cells through the wrapper; that hop is gone,
but the per-cell check stays, because the public getters convert more permissively than
the mapper's own rules and the check keeps both paths' results identical.

## Remaining costs and limits

Parity is incomplete. XLSX reads remain about 56% slower than ExcelReader.NET, XLSX
typed parsing 40%, XLSB reads 28% and XLSB typed parsing 28%. The remaining read
cost is structural rather than one hidden hot spot. OfficeIMO proves the whole
part before delivering the first row and gives the data reader a fixed field
count. It therefore makes a validation and indexing pass and then a delivery
pass; ExcelReader.NET reads each row once. Inflation takes about a fifth of the
XLSX and XLSB reads. Typed parsing also qualifies every mapped cell through the
data-reader wrapper before the compiled mapper reads it.

The XLS and XLSB writers are the `WriteData` plus `Save` route. Before it saves,
that route retains every source row and boxed selector value, so the generated
rows survive into older GC generations. On its own, the measured registration
takes about 7 ms of each write against 0.7 ms to generate the same rows. XLSB
deflate takes about another 5 ms; its BIFF12 payload is 5.7 MB against 5.1 MB from
ExcelReader.NET; the difference matches writing the three numeric columns as
8-byte reals instead of 4-byte RK values. XLS also
builds its compound file as one more complete copy. Removing the buffering would
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
