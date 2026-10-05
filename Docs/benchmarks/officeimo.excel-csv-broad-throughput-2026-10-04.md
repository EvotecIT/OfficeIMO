# Excel and CSV read, copy and save allocation — 2026-10-04

Commit `8b84a1676` reduces managed allocation in CSV incremental reads, quoted
row materialization, XLSX rich shared strings, prefixed worksheets, and package
copies. The baseline is `0d0d5a0b6`, after the
[earlier small-workload changes](officeimo.excel-csv-throughput-2026-10-04.md).

The investigation covers multiple API contracts and data sizes. It does not
establish a general throughput lead or portable performance budgets. Other
applications and builds remained active on the workstation; timing differences
for several unchanged workloads reversed between processor groups.

## Allocation results

These are warmed BenchmarkDotNet managed allocations per completed operation.
They do not measure live retained objects, pooled capacity, native memory, or
peak process memory. MiB and KiB use powers of 1,024.

| Workload | Before | After | Reduction |
| --- | ---: | ---: | ---: |
| CSV, 100,000 plain rows, incremental async scan | 48.91 MiB | 24.30 MiB | 50.3% |
| CSV, 100,000 multiline rows, incremental async scan | 176.76 MiB | 57.79 MiB | 67.3% |
| CSV, 100,000 multiline rows, snapshot async scan | 207.29 MiB | 174.48 MiB | 15.8% |
| CSV, 25,000 wide rows, incremental async scan | 66.73 MiB | 57.28 MiB | 14.2% |
| CSV, 25,000 quoted rows, independent streaming rows | 42.94 MiB | 32.25 MiB | 24.9% |
| CSV, 25,000 quoted rows, reusable callback | 31.11 MiB | 20.43 MiB | 34.3% |
| XLSX, 25,000 rich shared strings, ordinary worksheet | 5.34 MiB | 3.05 MiB | 42.8% |
| XLSX, 25,000 ASCII shared strings, prefixed worksheet | 7.64 MiB | 1.37 MiB | 82.1% |
| XLSX, 25,000 rich shared strings, prefixed worksheet | 11.64 MiB | 3.08 MiB | 73.5% |
| XLSX, 2,500-row package copy and save | 157.97 MiB | 53.23 MiB | 66.3% |
| XLSX, 25,000-row package copy and save | 1,615.89 MiB | 510.12 MiB | 68.4% |

Unchanged result contracts remain visible in the evidence. Plain snapshot CSV
reads allocate about 71.01 MiB at 100,000 rows; wide snapshot reads allocate
about 103.86 MiB at 25,000 rows. Ordinary ASCII, Unicode, and entity-tail XLSX
reads have effectively unchanged allocation. A 100-row package copy remains
about 1.91 MiB, and a 25,000-row values-only copy remains about 443.54 MiB.

The first 2,500-row values-only copy comparison measured 39.64 → 45.23 MiB.
A longer warmup measured 45.185 → 45.184 MiB. Both observations are retained;
the follow-up does not reproduce an allocation regression. Package copying
and values-only copying have different contracts and are not ranked against
one another as equivalent operations.

## Changes and correctness boundaries

The incremental CSV reader reuses delimiter text, field storage, and bounded
line/record builders. Header and inference samples keep independent arrays;
values returned to callers do not change when reading the next row. The ordinary
quoted-row parser validates the closing-quote grammar before creating the final
string directly. Multicharacter delimiters and unsupported quoted forms retain
their established parser paths.

Shared-string loading reuses a builder between rich-text items while preserving
each completed string and excluding phonetic runs. Builders above the retention
threshold are discarded. The indexed worksheet reader also accepts a consistent
element prefix. Prefixed documents pass complete XML and resolved-namespace
validation before indexed rows are exposed; unsupported prefix layouts use the
existing fallback.

Adjacent reader tests exposed two correctness defects. Literal CR/CRLF
normalization now occurs before entity expansion, preserving referenced carriage
returns. Reading an empty inline-string element advances the XML reader instead
of leaving the caller on the same node indefinitely. The new fixtures cover
ordinary and prefixed elements, mixed-prefix fallback, namespace rebinding,
malformed trailing XML, empty cells, formulas, booleans, Unicode, and line endings.

Package copying replaces shared-string value nodes directly and avoids saving
unchanged reference rewrites. Saving simple inline strings avoids constructing
each cell's complete XML through `OuterXml`. Rich or extended markup retains
the SDK serializer. XML character validation, whitespace attributes, CR/CRLF,
rich runs, and reopened values remain covered by artifact tests.

Public APIs and production dependencies are unchanged.

## Workloads and timing controls

The initial diagnostic matrix contains 114 CSV and 119 Excel cases: synchronous
and asynchronous reads, first-row and full scans, typed mapping, span and string
results, wide records, text/file/DataReader writes, shared strings, public
parallel requests, DataTable/object reads, package writes, round trips, copies,
and native binary formats. Synthetic workloads complement the hash-pinned 65K
corpus. Comparison adapters consume equivalent fields and validate complete
output; snapshot and incremental contracts remain separate.

The retained before/after packet contains 130 native cases, including rejected
experiments, and 3,168 rotated samples. PowerForge rotates engine order and
retains every measured sample. An identical-baseline control precedes the
candidate comparisons. Runs use Normal priority and both verified 96 MiB L3
processor groups on the Ryzen 9 9950X3D2, with masks `0xFFFF` and `0xFFFF0000`.
Those masks are specific to this workstation.

Across the five 25,000-row prefixed-worksheet text shapes, the rotated candidate
medians are 0.41–0.54 times baseline on the first processor group and 0.43–0.50
on the second. Allocation falls in all ten prefixed cases, including 256-row
inputs. Ordinary worksheet timing is less consistent and does not establish
a general speedup.

Longer rotated follow-ups do not reproduce the initial slowdowns in multiline
streaming rows, wide incremental reads, or 2,500-row values-only copies. Wide
snapshot reading is within 2% on the first group but 9.6% slower on the second;
that timing signal remains unqualified. Its allocation is unchanged.

A buffered-encoding experiment improved one tabular export but slowed several
short- and long-text workloads without reducing allocation. It was removed.
Its complete 13-workload before/after result is retained with the rejected
stage label, rather than counted as an improvement.

## Rejected multicharacter CSV writer batching

A second writer experiment batched ordinary `IDataReader` text-delimiter rows
through the existing formatted-row buffer. It preserves correctness across
Windows .NET 10 and .NET 8 (634 tests each), .NET Framework 4.7.2 (444 tests),
and Linux .NET 10 under WSL (634 tests). Independent read-only review reports
no actionable defects. Those checks do not establish a performance benefit.

The output-validated matrix covers 22 workloads: four file-output text shapes,
short and long in-memory Notes/JSON/quote-heavy fields, both quote modes, and
unchanged comma-delimiter controls. It retains 44 native cases and 2,112 rotated
samples across both processor groups. Native measurements use 12 warmups,
12 measurements and eight operations per measurement. Rotated runs use
12 warmups, 24 measurements and four operations per sample, at Normal priority
with all outliers retained.

Short Unicode file-output medians improve about 6% with `Always` quoting, but
become 4–6% slower with `AsNeeded`. Short JSON in-memory output allocates another
6–8 KiB per operation. Other timings are mixed, and unchanged controls vary too.
The production change is removed because these results do not justify enabling
it generally. The expanded delimiter benchmarks and completed-row cancellation
and error contracts remain; the final suite against restored production code
passes all 634 .NET 10 tests.

The complete [native measurements](excel-csv-broad-throughput-2026-10-04/rejected-csv-batching-native.json),
[rotated samples](excel-csv-broad-throughput-2026-10-04/rejected-csv-batching-rotated.json),
and [provenance with the rejected patch](excel-csv-broad-throughput-2026-10-04/rejected-csv-batching-provenance.json)
record favorable and unfavorable cases. To reproduce the experiment, build the
recorded coverage commit as the baseline, then apply the embedded production
patch in a separate candidate checkout. Run the recorded cases through the
snapshot comparison runner described below, using the same benchmark assembly
for both snapshots. The rejected candidate is not part of the current product.

An additional single-character `Always`-quoting batch experiment is also removed.
Its eight-workload screen retains 16 native cases: short notes and JSON, long
JSON, ASCII file output, and unchanged `AsNeeded` controls. Short always-quoted
JSON adds 7,728 allocated bytes per operation, and both short always-quoted text
cases measure slower. Control timings also vary substantially, so the screen
does not establish stable speed ratios. It gives no basis for adopting the
change. All 638 .NET 10 CSV tests pass before and after removal; the additional
always-quoted cancellation and formatting-failure contracts remain. The
[screen packet](excel-csv-broad-throughput-2026-10-04/rejected-csv-always-quoted-batching.json)
includes complete observations, rejected source, binary hashes and case settings.
No rotated, other-runtime, or independent-review qualification is claimed for
this rejected experiment.

## Rejected direct destination reuse for CSV DataReader writes

An experiment writes default `IDataReader` records directly into an exact
`StringWriter` destination's existing builder. The initial 12-workload screen
shows lower quote-heavy medians, but those gains do not hold consistently in the
broader comparison. The production change is removed.

The full matrix covers 22 workloads: short/long notes, JSON and dense quotes;
25,000/100,000-row mixed, quoted and multiline typed tables; public parallel
requests; file output; and unchanged quote-mode and delimiter paths. All 44
native cases and 2,112 rotated samples pass complete-output validation. The
100,000-row quoted typed writer is about 17% slower on both processor groups,
while allocation falls only from 29,395,312 to 29,351,872 bytes. Short-field and
long-JSON timing changes reverse between groups. Unchanged file and delimiter
controls also vary, so the experiment does not establish stable gains elsewhere.

The initial screen's pooled-buffer allocation differs from the full native run:
several baseline cases add 16,390 bytes per operation in the latter. The packet
retains both observations rather than treating one warmed pool state as a stable
memory budget. Parallel-request cases can reach the sequential fallback and are
not unchanged-code controls despite their diagnostic case labels.

Candidate correctness passes 639 CSV tests on Windows .NET 10/.NET 8, 449 on
.NET Framework 4.7.2, 639 on Linux/WSL .NET 10, and 639 each on macOS ARM64
.NET 10/.NET 8. The `netstandard2.0` build succeeds and independent review reports
no actionable findings. Correctness alone does not justify the performance tradeoff.
The completed-row regression test remains: a reused writer preserves its existing
text, prior successful calls and completed rows when a later row fails, and remains
usable afterward. Restored production code passes all 639 Windows .NET 10 tests
and 449 .NET Framework tests.

The [native screen and full matrix](excel-csv-broad-throughput-2026-10-04/rejected-csv-direct-buffer-native.json),
[rotated observations](excel-csv-broad-throughput-2026-10-04/rejected-csv-direct-buffer-rotated.json),
and [source, binary, runtime and reproduction packet](excel-csv-broad-throughput-2026-10-04/rejected-csv-direct-buffer-provenance.json)
retain the rejected patch and every measured case. Timing remains specific to
this busy workstation; managed allocation does not measure retained or peak memory.

## Peer refresh and remaining gaps

The [post-change peer packet](excel-csv-broad-throughput-2026-10-04/peer-refresh-native.json)
retains 109 CSV and 26 Excel native cases, including each observation, allocation
result and job setting. It covers async read contracts, mapping, typed scans,
text/file/DataReader writing and the 65K corpus. The CSV production source is
unchanged between the broad allocation commit and `882226a45`. Later large XML
reader changes supersede the Excel read results; use their separate evidence
before assessing the current reader. These peer reports lack independent
per-run binary fingerprints and are diagnostic evidence, not a final ranking.

The refresh still exposes short-JSON writing costs. At 1,000 short JSON rows,
OfficeIMO measures about 0.50 ms with `AsNeeded` and 0.47 ms with `Always`,
versus 0.28 and 0.27 ms for CsvHelper. OfficeIMO allocates about 192 and 203 KiB,
versus 618 and 651 KiB. Long JSON and dense-quote results favor OfficeIMO in
this run. Those gains do not close the short-field throughput gap, and timing
differs substantially from the earlier small-workload measurements.

For the equivalent 25,000-row compact XLSX export, OfficeIMO measures
19.11/20.38 ms across the two processor groups, SpreadCheetah 14.03/14.13 ms,
LargeXlsx 16.32/15.90 ms and Sylvan 18.96/18.92 ms. Every implementation's
headers and cells are independently reopened and checked. These runs retain
all outliers and use fixed placement, but do not rotate cross-engine order;
the busy-host timing limits still apply. General export throughput remains
an open target.

## macOS conformance

The [macOS evidence packet](excel-csv-broad-throughput-2026-10-04/macos-conformance.json)
records the integrated source at `a97198615` on Apple M4, macOS 27.0.1 and
.NET SDK 10.0.112. The clean detached worktree passes all 638 CSV tests and
5,376 Excel tests, with five existing Excel skips. Both comparison projects
build for .NET 10.

BenchmarkDotNet Dry runs complete 133 CSV and 22 Excel comparison cases with
their existing complete-output setup validators. The CSV cases cover text,
file and DataReader writing, async reads, wide async reads, typed scans,
automatic mapping, custom delimiters and the 65K corpus. Excel covers compact
DataReader export, 25,000/250,000/1,000,000-row complete reads and the 65K
corpus. The packet retains each case, report hash, binary hash and test result.

These runs establish conformance on ARM64. Dry timing and allocation are not
ranked, and concurrent CI and simulator activity prevents a quiet-host timing
claim. Portable throughput and retained/peak-memory budgets remain open.

## Million-row memory and output checks

The separate CSV memory lane measures 100,000 and 1,000,000 plain/multiline rows,
both reader modes, and first-row/full-scan operations. All 96 measured samples
pass their output checks. PowerForge samples managed and resident memory every
5 ms; these observations are lower bounds, not exact peaks or retained-heap
measurements. Instrumentation dominates the very short incremental first-row
operation, so these timings do not establish a portable first-row latency.

At one million rows, incremental first-row managed increase remains about
0.04–0.05 MiB. Full incremental scans show roughly 45 MiB sampled managed
increase before and after. Snapshot scans remain hundreds of MiB. The lower
allocation figures above therefore do not imply a measured peak-memory reduction.

Synchronous and asynchronous XLSX exports also pass a one-million-row,
four-column validation for each compared writer. Reopening checks every header,
typed value, row count, and the single-sheet result. This is large-output and
reader-fallback correctness evidence, not a throughput or memory ranking.

## Default typed formatting during asynchronous document saves

`CsvDocument.SaveAsync` reuses the existing buffer formatter for default quoting,
null/date handling and formula preservation with a single-character delimiter.
Typed numbers can be appended without first creating a formatted string. Custom
options use the established serializer, and record emission retains the existing
cancellation, asynchronous output, compression and caller-ownership behavior.

Qualification covers 72 workloads: 1,000 and 25,000 rows across plain text,
quoted Unicode/multiline text and mixed JSON/typed values with no compression,
GZip, Deflate, Brotli and ZLib; 100,000-row cases cover all three shapes with no
compression and GZip. Setup compares every output field against an independently
formatted reference. The complete comparison contains 144 native cases and
6,912 rotated samples. Each rotated process holds one six-workload batch.

The following table shows candidate/baseline median time on the two confirmed
96 MiB L3 cache domains. Ratios below one are faster. Managed allocation falls
by approximately 1.71 MiB per 25,000-row async save and 6.86 MiB per 100,000-row
async save in the native comparison.

| Async save workload | Rows | Domain A, `0xFFFF` | Domain B, `0xFFFF0000` |
|---|---:|---:|---:|
| Plain, uncompressed | 25,000 | 0.645 | 0.698 |
| Quoted, uncompressed | 25,000 | 0.726 | 0.773 |
| Mixed JSON, uncompressed | 25,000 | 0.796 | 0.747 |
| Quoted, GZip | 25,000 | 0.813 | 0.777 |
| Plain, uncompressed | 100,000 | 0.686 | 0.615 |
| Quoted, uncompressed | 100,000 | 0.699 | 0.620 |
| Mixed JSON, uncompressed | 100,000 | 0.601 | 0.535 |
| Quoted, GZip | 100,000 | 0.727 | 0.759 |

Small cases remain timing-sensitive. Four initial 1,000-row async cases on
domain B have median ratios of 1.10–1.56. A bounded follow-up with longer warmup,
16 operations per sample and 48 measured samples has ratios of 0.60–0.69 on
domain A and 0.56–0.69 on domain B. Identical-binary controls in that follow-up
have ratios of 1.03–1.27, showing substantial host/runtime variation. The packet
retains both observations. An earlier screen also shows an unexpected difference
in synchronous source that disappears substantially with tiered compilation
disabled; synchronous gains are not attributed to this change.

Correctness passes 655 CSV tests on Windows .NET 10 and .NET 8, 463 on .NET
Framework 4.7.2, 655 on Linux/WSL .NET 10, and 655 on each of macOS ARM64 .NET 10
and .NET 8. The `netstandard2.0` product build has no warnings or errors. Ten
formatting parity cases cover typed values, cultures, delimiters, quoting,
custom value policies, formulas and custom span-formatting fallback. One fresh
read-only review reports no actionable findings. Timing is Windows .NET 10
evidence; other platforms and targets provide correctness proof. Warmed managed
allocation does not establish retained or peak memory.

The packet retains [native measurements](excel-csv-broad-throughput-2026-10-04/async-document-formatting-native.json),
[full rotations](excel-csv-broad-throughput-2026-10-04/async-document-formatting-rotated.json),
[screening and follow-up controls](excel-csv-broad-throughput-2026-10-04/async-document-formatting-diagnostics.json),
and [source, binary, runtime, review and reproduction details](excel-csv-broad-throughput-2026-10-04/async-document-formatting-provenance.json).
These measurements do not establish a cross-library ranking or close the
remaining spreadsheet throughput and portable-memory targets.

## Smaller CSV writer buffer experiments

The established 256 KiB write buffer remains the default. Two experiments
separate write-buffer sizing from read-buffer sizing and try 16 KiB and 64 KiB
across document saves, byte serialization, file row writers and sequential or
parallel DataReader exports. Both reduce fixed managed allocation, but neither
qualifies as a universal default.

The 16 KiB screen saves approximately 0.94–1.41 MiB per operation. Its rotated
25,000-row file row-writer median ratios are 1.161 and 1.136 on the two CPU
domains; asynchronous plain-file saves have ratios of 1.141 and 1.253. These
repeatable file throughput costs outweigh the allocation benefit for a general
policy.

The 64 KiB screen saves approximately 0.75–1.13 MiB per operation. Some stream
and compressed workloads improve, but native and rotated timings disagree.
Longer identical-build controls show 10–21% variation between instances in the
1,000-row file row-writer case. The candidate follow-up changes direction on
that case between domains. Larger file workloads still show material costs on
one domain: the 25,000-row asynchronous plain-file save has median ratios of
approximately 1.03 and 1.09, with mean ratios of 1.03 and 1.15. This evidence
does not establish a dependable throughput improvement across the writer APIs.

The packet retains 220 native cases, including the initial baseline, and 6,528
rotated samples across both experiments and the 64 KiB controls. Every measured
sample succeeds. Each candidate passes 655 CSV tests on Windows .NET 10 and
.NET 8 and 463 on .NET Framework 4.7.2. Setup validates all eight API outputs
across 39 fixture combinations for each actual runtime and snapshot build,
checking complete decoded CSV and every field. The broader 240-workload timing
matrix is not run after these candidates fail screening qualification.

The [native results](excel-csv-broad-throughput-2026-10-04/writer-buffer-native.json),
[rotations and controls](excel-csv-broad-throughput-2026-10-04/writer-buffer-rotated.json),
and [patches, binary manifests, dispositions and reproduction details](excel-csv-broad-throughput-2026-10-04/writer-buffer-provenance.json)
preserve favorable, unfavorable and conflicting observations. These are
Windows workstation experiments. Managed allocated bytes do not establish
retained or peak memory; file timings include operating-system caching and
exclude durable-storage flushes. Neither experiment changes the product.

## Serialized CSV stream saves without a second byte array

`CsvDocument.Save(Stream)` transfers the completed serialization buffer directly
through the shared stream writer. It retains staging: formatting, encoding and
compression finish before the destination is touched. Seekable destinations are
truncated before byte emission and rewound after success; forward-only streams
receive bytes at their current position. Caller streams remain open. `ToBytes`
and the fixed-capacity `ToStream` contract retain their independent byte arrays.

The matrix covers 39 stream-save workloads and six adjacent controls. It includes
1,000 and 25,000 plain, quoted and mixed JSON rows with no compression, GZip,
Deflate, Brotli and ZLib; 100,000 rows with no compression or GZip; one plain row;
and three long Unicode rows with no compression or GZip. Controls exercise
`ToBytes`, asynchronous saves, path saves, DataReader writes and file row writers.
All 39 target workloads allocate less on actual .NET 10 and .NET 8 runtimes.
At 100,000 rows, warmed .NET 10 allocation per operation is:

| Shape and compression | Before | After | Reduction |
| --- | ---: | ---: | ---: |
| Plain, none | 34.334 MiB | 29.096 MiB | 5.238 MiB |
| Quoted, none | 37.831 MiB | 31.466 MiB | 6.365 MiB |
| Mixed JSON, none | 61.985 MiB | 52.412 MiB | 9.573 MiB |
| Plain, GZip | 14.719 MiB | 13.409 MiB | 1.309 MiB |
| Quoted, GZip | 14.810 MiB | 13.458 MiB | 1.353 MiB |
| Mixed JSON, GZip | 14.894 MiB | 13.500 MiB | 1.394 MiB |

The qualification retains 180 native cases and 8,928 rotated samples, including
identical-build controls and longer follow-ups on both processor groups. The
initial 25,000-row plain/Brotli medians are 1.129/1.073 times baseline; the longer
follow-up measures 0.986/1.011. Short-case timing remains inconclusive. For one
plain row, follow-up medians are 0.980/1.264 with means 0.599/0.836; an unchanged
async control also moves to 0.952/1.136. Identical-build long-Unicode controls
measure 1.720/1.861. These observations do not establish a general speedup or
portable latency budgets. The change is accepted for its allocation reduction.

Windows CSV correctness passes 666 tests on each modern runtime and 472 on
.NET Framework 4.7.2. Linux/WSL passes 666 on .NET 10; macOS ARM64 passes 666 on
each modern runtime. The product builds for `netstandard2.0` without warnings or
errors. Existing Core stream contracts pass three focused linked tests on each
Windows runtime; this does not claim the complete Shared.Tests suite. Independent
decoded-text and field checks validate 1,248 outputs across eight APIs, two builds
and two actual runtimes.

One independent read-only review finds a failure-ordering regression in the first
candidate: a fixed-length mapped stream is overwritten before resizing fails.
The final shared buffer-segment transfer rejects resizing before emitting bytes.
A real mapped-file regression fails on that candidate, passes with the baseline
CSV binary, and passes with the fix. The targeted review confirmation reports no
additional findings. Serialization failures after large prior rows preserve the
destination's bytes and position with UTF-8, UTF-16 and compression.

The [native observations](excel-csv-broad-throughput-2026-10-04/stream-save-native.json),
[rotations and controls](excel-csv-broad-throughput-2026-10-04/stream-save-rotated.json),
and [source, binary, validation and reproduction packet](excel-csv-broad-throughput-2026-10-04/stream-save-provenance.json)
retain all outcomes. Staging, destination capacity and writer buffers remain;
managed allocation does not measure retained or peak memory. General throughput,
small-workload timing and portable memory qualification remain open.

## Validation and reproduction

- Windows CSV correctness: 630 tests on .NET 10, 630 on .NET 8, and 440 on .NET Framework 4.7.2.
- Windows Excel: the complete .NET 10 non-performance suite passes 5,225 tests, with five existing opt-in/platform skips. Focused .NET 8 and .NET Framework suites pass 355 and 350 tests.
- Ubuntu 24.04 under WSL: 630 CSV and 362 focused Excel tests pass on .NET 10.0.12. This is WSL correctness evidence, not native-Linux timing qualification.
- Both benchmark projects build on .NET 8 and .NET 10; both product projects build for `netstandard2.0` without warnings or errors.
- Two independent read-only reviews cover the separate CSV/SST/copy and prefixed-reader changes. Neither reports actionable defects. The reviewed patches match the committed source.

The packet retains [native measurements](excel-csv-broad-throughput-2026-10-04/native.json),
[rotated measurements and controls](excel-csv-broad-throughput-2026-10-04/rotated.json),
[sampled memory](excel-csv-broad-throughput-2026-10-04/memory.json),
[provenance](excel-csv-broad-throughput-2026-10-04/provenance.json), and
[validation](excel-csv-broad-throughput-2026-10-04/validation.json).
The [reproduction guide](excel-csv-broad-throughput-2026-10-04/reproduction/README.md)
describes the snapshots, output checks, stage labels, and runner requirements.

The [product roadmap](../ROADMAP.md#spreadsheet-and-csv-delivery-order) retains
the remaining work: large general XLSX read/export throughput, portable
first-row latency and retained/peak-memory budgets, quieter-host confirmation,
and native Linux/macOS measurements. These results do not close those targets.
