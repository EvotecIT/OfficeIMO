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

## CSV document text allocation — 2026-10-05

`CsvDocument.ToString()` uses the existing row formatter to append default typed
values directly to its `StringWriter` buffer. On modern runtimes this avoids
intermediate scalar strings. Custom date/null formatting, UTC conversion,
formula escaping, selected quoting and multicharacter delimiters retain their
existing paths. Public APIs and production dependencies are unchanged.

Windows and macOS native measurements use the same owner source and benchmark
assembly within each before/after pair, on actual .NET 10 and .NET 8. These
Windows .NET 10 figures measure managed allocation per complete text result:

| Document | Before | After | Reduction |
| --- | ---: | ---: | ---: |
| 25,000 plain rows | 6.83 MiB | 5.13 MiB | 25.0% |
| 100,000 plain rows | 27.86 MiB | 21.01 MiB | 24.6% |
| 100,000 quoted Unicode rows | 30.65 MiB | 23.79 MiB | 22.4% |
| 100,000 mixed JSON rows | 42.12 MiB | 35.26 MiB | 16.3% |

All three 25,000-row shapes save about 1.71 MiB; all three 100,000-row shapes
save about 6.86 MiB on both platforms and runtimes. The 1,000-row cases save
about 5.6 KiB. Allocation does not improve everywhere: a one-row document adds
24 bytes, three long Unicode rows add 72 bytes, and the 64-row long-text delta
ranges from about 2 KiB lower on Windows .NET 10 to 1.5 KiB higher on .NET 8
and macOS. Allocation figures do not measure retained or peak memory.

The initial matrix has 14 text cases and five unchanged save/export controls.
Its three one-row labels share a null Notes field and therefore the same
payload. They remain in the evidence as null-field controls. A corrected
fixture retains its text when there is only one row; four additional profiles
exercise plain text, quoted Unicode, JSON and a long Unicode field. Their
four output hashes differ and match across both builds, platforms and runtimes.
Larger fixtures retain their original null-field coverage and payloads.

The packet retains 184 native cases and 9,840 rotated samples, including every
identical-baseline control and unfavorable measurement. Windows runs separate
both verified processor groups at Normal priority. macOS uses operating-system
scheduling on the M4; its CPU domains are not fixed. Native runs use 24 warmups,
12 measurements and four operations per measurement, with eight operations for
the corrected one-row profiles. Full rotated runs use 48 warmups, 24 measurements
and four operations; follow-ups use 64 warmups, 48 measurements and eight
operations. All outliers remain, and no output validation failed.

The 25,000-row plain-text rotated medians are 0.69 and 0.72 times baseline on
the Windows groups and 0.61 on macOS. Longer 100,000-row quoted-text follow-ups
measure 0.69, 0.67 and 0.68 respectively. These are workload-specific observations.
The 100,000-row mixed-JSON median does not improve on the second Windows group,
and small/long-text timing is mixed. Identical-baseline and unchanged-control
variation prevents portable latency budgets or a general throughput claim.
The change is accepted for its large typed-document allocation reduction.

Correctness passes all 675 CSV tests on Windows .NET 10/.NET 8, 481 on .NET
Framework 4.7.2, 675 on Linux/WSL .NET 10, and 675 on each macOS runtime. Focused
literal fixtures cover quoting, Unicode, culture, nulls, custom formatting and
formatter failures. The .NET Standard 2.0 build passes. Independent read-only
review reports no actionable defects; its additional formatter-proof gaps are
covered by the final tests.

The [native measurements](excel-csv-broad-throughput-2026-10-04/csv-text-native.json),
[rotated samples](excel-csv-broad-throughput-2026-10-04/csv-text-rotated.json),
[portable output contracts](excel-csv-broad-throughput-2026-10-04/csv-text-output-contracts.json),
[corrected one-row packet](excel-csv-broad-throughput-2026-10-04/csv-text-one-row.json),
and [source, binary and execution provenance](excel-csv-broad-throughput-2026-10-04/csv-text-provenance.json)
retain the full matrix. The provenance records the excluded run whose source
fingerprint changed during another benchmark edit; it contributes no samples.
The [writer and DataTable diagnostic profiles](excel-csv-broad-throughput-2026-10-04/writer-datatable-profiles.json)
retain sampled allocation and CPU attribution. Those traces include setup and
validation; they are not exact per-operation allocation or peak-memory evidence.

## Late repeated worksheet rows — 2026-10-05

Commit `048b089925` preserves updates from late physical worksheet rows, including
rows that appear after the requested range has already been encountered. Array,
row, column, chunk, dictionary, typed-object, DataTable and DataReader paths keep
the later values. Readers qualify row order before publishing results that could
otherwise become stale. The unsorted typed-stream fallback buffers pending row
fragments within its existing resource limit.

Regression fixtures cover UTF-8 and UTF-16, small and large ranges, sequential
and parallel entrypoints, cancellation-capable paths, and the direct-schema
DataReader. The final reader selection passes 742 tests on Windows .NET 10/.NET
8, Linux/WSL .NET 10 and both macOS runtimes; .NET Framework 4.7.2 passes 738.
Complete Windows Excel suites pass 5,464 tests with five existing skips on each
modern runtime. The .NET Standard 2.0 build passes. Independent review exposed
three additional reachable variants; all were reproduced and fixed, and the
targeted confirmation reports no remaining actionable findings.

The correctness cost is retained separately from subsequent optimization work.
The 39-workload native comparison contains 156 before/after observations across
Windows .NET 10 and .NET 8. Both sides use the same benchmark assembly; only the
Excel library and symbols differ. Each observation retains all 12 measurements
after 24 warmups, four invocations per measurement, Normal priority and the
verified `0xFFFF` processor group. Complete-output setup validation passes every
workload on both runtimes.

Short-prefix reads now traverse the remaining physical rows to preserve late
updates. For a 100-row prefix of a 25,000-row worksheet, native mean ratios are
11.60 and 11.97 on .NET 10 without/with column inference, and 6.25 and 10.21 on
.NET 8. The 2,500-row prefix ratios are 1.64/1.65 and 1.98/1.85 respectively.
These are cost observations on this host with fixed execution order, not
portable latency budgets. A performance candidate must be compared with this
corrected baseline; the former early exits cannot serve as a valid target.

The [qualification and complete native-cost packet](excel-csv-broad-throughput-2026-10-04/xlsx-reader-termination-qualified.json)
retains the source and binary fingerprints, output proof, test counters, review
boundary, measurements and unfavorable ratios. Wide unsorted typed-stream
retained/peak memory and first-row latency remain open measurements.

## Public XLSX readers through retained packages — 2026-10-05

Commit `160bd1049` lets qualified, ordered XLSX worksheets continue through the
existing XML data reader using the retained package stream when indexing is
unavailable. Complete worksheet validation establishes row order and used
bounds before values are exposed. Unsupported culture, options, worksheet
structures and repeated or unsorted rows retain their existing fallback paths.
The change introduces no public API or production dependency.

The comparison uses the same benchmark assembly and dependencies on both sides;
only the Excel library and symbols differ. Setup validates all four fields of
every row, headers, row count and the aggregate observation. The peer runs also
require identical decoded worksheet, style and shared-string parts. Rotated
runs use eight warmups and retain all 12 measurements without removing outliers.
The PowerShell observations run on .NET 10; separate native jobs qualify .NET 8
and .NET 10. Windows measurements keep Normal priority and distinguish the two
processor groups. macOS measurements use the operating system's scheduling.

| Rows | Windows group 0 before → after, median ms | Windows group 1 before → after, median ms | macOS before → after, median ms |
|---|---:|---:|---:|
| 25,000 | 95.15 → 83.31 | 58.44 → 55.43 | 53.19 → 49.22 |
| 250,000 | 790.43 → 845.42 | 678.09 → 671.83 | 503.05 → 470.25 |
| 1,000,000 | 3,500.83 → 3,233.81 | 2,339.06 → 2,542.56 | 1,870.65 → 1,890.37 |

Managed allocation falls at every size in these rotated comparisons: by about
203–224 KB per operation. Throughput is mixed. The complete native jobs retain
additional unfavorable observations: the Windows .NET 8 million-row median is
3,091.35 → 3,848.71 ms, and macOS .NET 10 is 1,630.07 → 1,859.57 ms. Conversely,
the macOS .NET 8 million-row median is 3,214.94 → 2,057.28 ms. These observations
remain part of the evidence; they do not establish a portable throughput win.

The separate held-reader runs open a reader, validate the first row's four
fields, then measure managed memory after GC while retaining that reader. Each
size uses six measurements per side in a fresh worker. On Windows, held managed
increases fall from 60,884 → 23,177 bytes at 25,000 rows and approximately
69,300 → 30,309 bytes at the larger sizes. Those are increases over a warmed
baseline that already contains pooled buffers, rather than the total pool
footprint. Sampled managed and resident peaks are lower bounds. Timings include
GC and sampling and are not throughput observations.

Equivalent full four-field peer scans still show a material throughput gap:

| Host and rows | OfficeIMO median ms | Sylvan median ms | ExcelReader.NET median ms |
|---|---:|---:|---:|
| Windows group 0, 250,000 | 724.08 | 309.32 | 91.82 |
| Windows group 1, 250,000 | 696.03 | 309.05 | 88.74 |
| macOS, 250,000 | 778.08 | 366.91 | 113.78 |
| Windows group 0, 1,000,000 | 3,092.74 | 1,420.00 | 405.64 |
| Windows group 1, 1,000,000 | 2,970.97 | 1,379.03 | 403.61 |
| macOS, 1,000,000 | 3,034.51 | 1,363.46 | 419.02 |

These warmed .NET 10 runs use Sylvan 0.5.8 and ExcelReader.NET 5.1.1. Every
implementation consumes the same four fields and passes complete setup
validation. First-row methods have different eager-validation contracts and
are excluded from this peer ranking. The allocation improvements do not close
the large-scan throughput target.

The [Windows native, memory and peer packet](excel-csv-broad-throughput-2026-10-04/xlsx-public-reader-stream-windows.json),
[macOS packet](excel-csv-broad-throughput-2026-10-04/xlsx-public-reader-stream-macos.json),
[rotated before/after observations](excel-csv-broad-throughput-2026-10-04/xlsx-public-reader-stream-rotated.json)
and [second Windows processor placement](excel-csv-broad-throughput-2026-10-04/xlsx-public-reader-peers-second-placement.json)
retain raw samples, source and binary fingerprints, runtime provenance and
decoded input proof. Native method order differs from the rotated runs; the
packets preserve both results.

## Typed reader integration and fallback contracts — 2026-10-05

Typed readers merge later physical row fragments before assigning properties.
Omitted cells preserve initialized values; present nulls or blanks follow the
existing conversion rules. Presence metadata remains internal to the typed
mapper. Ordinary public data readers retain their existing signatures and
DBNull behavior. DataTable row reuse preserves earlier values when a later
fragment omits a cell and retains conservative blank normalization for wide
rows.

Commit `48f77cd75` preserves typed fallback when worksheet indexing encounters
an uncached shared-formula follower outside the requested columns. The typed
mapper declines that source before header binding, object construction or
property assignment. Its existing XML reader handles the narrower projection.
Invalid styles and shared-string references still fail, and public data readers
retain their explicit rejection of unresolved shared formulas. Cancellation
before index ownership transfers returns the candidate's pooled buffers.

The shared-formula regression is reproduced in ten of sixteen small and large
typed-read cases before the fix. All sixteen pass afterward; four additional
cases preserve invalid-reference failures. Coverage includes Automatic,
Sequential, Parallel and streaming APIs and both cached-result options.
Windows and macOS full Excel suites pass 5,614 tests with five existing opt-in
skips on each modern runtime. Windows focused validation passes 994 tests on each modern runtime
and 987 on .NET Framework 4.7.2; the .NET Standard 2.0 build succeeds. A fresh
read-only integration review and its single targeted confirmation accept the
corrected ownership and fallback contracts.

The [Windows integration qualification](excel-csv-broad-throughput-2026-10-04/xlsx-reader-integration-windows.json)
and [macOS qualification](excel-csv-broad-throughput-2026-10-04/xlsx-reader-integration-macos.json)
retain source fingerprints, test counts, TRX hashes and review scope. Correctness
qualification does not close the ordered wide UTF-16 typed-read regressions,
macOS small sequential timing signal, or the remaining DataTable .NET 8 timing
controls.

## Rejected short quoted-field chunking — 2026-10-05

Routing short dense quoted text through the existing bounded chunk escaper is
rejected. Its 24 workloads cover 64/4,096-character notes, JSON and dense quotes,
AsNeeded/Always quoting, comma/multicharacter delimiters and 1,000-row public
DataReader writes. Setup validates every decoded field and requires identical
complete text from OfficeIMO and CsvHelper. Both comparison sides use the same
harness and dependencies.

The full comparison retains 288 native observations and 3,456 measurements:
.NET 8 and .NET 10, both Windows processor groups, and macOS. Each observation
uses eight warmups, 12 measurements, a 100 ms iteration target, an unroll factor
of one and no outlier removal. Short JSON is consistently slower on macOS by
approximately 18–35%; Windows observations are mostly 8–33% slower, with one
approximately 2% improvement. Allocation is essentially unchanged. Variation
in unchanged long-note controls is retained and is not attributed to the source
change.

The original source is restored on both hosts. The [Windows rejection packet](excel-csv-broad-throughput-2026-10-04/rejected-csv-short-quoted-windows.json)
and [macOS rejection packet](excel-csv-broad-throughput-2026-10-04/rejected-csv-short-quoted-macos.json)
retain complete reports, raw measurements and source, harness and dependency
fingerprints. The following peer qualification uses the restored baseline and
the same warmed policy.

## Refreshed quoted-text writer comparison — 2026-10-05

The restored writer is compared with CsvHelper 33.1.0 across the same 24 text
shapes, lengths, quote modes and delimiters on .NET 8 and .NET 10. Windows runs
each runtime on both processor groups; macOS uses operating-system scheduling.
The six suites retain 288 observations and 3,456 measurements. Each observation
uses eight warmups, twelve measurements, a 100 ms iteration target, Normal
process priority, an unroll factor of one and no outlier removal. Setup decodes
every field of all 1,000 rows and requires identical complete output from both
writers. Full parameter identities, rather than truncated display strings,
join the matched observations.

OfficeIMO has the lower median in 143 of 144 matched comparisons and lower
managed allocation in all 144. The remaining comparison is 64-character notes,
Always quoting and a comma delimiter on Windows .NET 8 group 0: 0.1052 ms versus
0.0997 ms, a ratio of 1.055. This negative result remains in the retained packet.

| Host and runtime | Lowest OfficeIMO/CsvHelper median ratio | Highest ratio | Cases with higher OfficeIMO allocation |
|---|---:|---:|---:|
| Windows group 0, .NET 10 | 0.019 | 0.913 | 0/24 |
| Windows group 1, .NET 10 | 0.015 | 0.930 | 0/24 |
| Windows group 0, .NET 8 | 0.011 | 1.055 | 0/24 |
| Windows group 1, .NET 8 | 0.016 | 0.861 | 0/24 |
| macOS, .NET 10 | 0.060 | 0.835 | 0/24 |
| macOS, .NET 8 | 0.045 | 0.981 | 0/24 |

Across the four short-JSON variants, OfficeIMO/CsvHelper median ratios range
from 0.195 to 0.620. Managed allocation is approximately 197–208 KB per operation
for OfficeIMO and 633–667 KB for CsvHelper. These measurements supersede the
earlier short-JSON timing signal for this exact warmed text-writer contract;
they do not identify the cause of the historical difference or qualify cold
startup, file output, mixed typed fields, reading, or other platforms.

The [Windows peer packet](excel-csv-broad-throughput-2026-10-04/csv-quoted-text-current-windows.json)
and [macOS peer packet](excel-csv-broad-throughput-2026-10-04/csv-quoted-text-current-macos.json)
retain complete reports, raw measurements, actual runtimes, source provenance
and binary fingerprints. All 91 measured CSV source files also match the
integrated CSV source. Rejected candidate snapshots are removed after their
reports, manifests and source patch are retained; baseline snapshots remain
available for further comparisons.

## Native index eligibility experiment — 2026-10-05

The V3 experiment checks for an early worksheet dimension before renting and
inflating a native worksheet buffer. A dimensionless 250,000-row, four-column
sheet exceeds the one-million-cell index boundary, so neither the initial
used-range index nor the later range index can use that buffer. The experiment
avoids the discarded allocation while retaining complete XML projection
validation. It remains an isolated experimental checkpoint; these results do
not establish performance acceptance for the integrated reader.

Both sides use the same harness and dependencies. Three dimensionless sizes
(25,000, 250,000 and 1,000,000 rows) and two successful-index controls (1,000 and
25,000 rows with dimensions) validate every field, header and row count. The
decoded worksheet, style and shared-string parts match. Rotated .NET 10 runs
retain twelve measurements per side on each Windows processor group and macOS;
separate native .NET 8 and .NET 10 jobs retain twelve measurements per side on
Windows group 0 and macOS.

Warmed allocation increases by approximately 13–30 KB per operation, and
timing results are mixed. The unfavorable native observations include the
following successful-index and streaming controls:

| Host and runtime | Case | Before median ms | V3 median ms |
|---|---|---:|---:|
| Windows group 0, .NET 10 | Indexed 25,000 | 39.81 | 70.35 |
| Windows group 0, .NET 10 | Dimensionless 250,000 | 545.00 | 818.61 |
| Windows group 0, .NET 8 | Dimensionless 1,000,000 | 2,455.47 | 3,327.02 |
| macOS, .NET 10 | Indexed 25,000 | 21.67 | 46.43 |
| macOS, .NET 10 | Dimensionless 1,000,000 | 1,732.77 | 1,862.26 |

Rotated runs retain their differing observations: the 250,000-row Windows
median changes by approximately +12% in group 0 and −9% in group 1, while macOS
changes by approximately −11%. These measurements do not justify attributing
all native timing differences to the header probe, discarding the negatives,
or accepting a portable throughput improvement.

A separate first-workbook measurement makes the discarded buffer visible.
Six fresh workers per side and case load the frozen assembly and options, then
perform their first OfficeIMO workbook read. Fixture production happens in
another process and validates every field; each worker validates the four
first-row fields before retaining the open reader during a forced-GC memory
observation. The operating system's file cache is warm. JIT and engine
initialization may contribute to elapsed time. Calling-thread allocation
excludes the memory sampler; sampled peaks are lower bounds. These are not
warmed throughput or equivalent first-row peer rankings.

| Host, dimensionless 250,000 rows | Before allocation bytes, mean | V3 allocation bytes, mean | Before held managed bytes, mean | V3 held managed bytes, mean | Before → V3 open/first-row ms, mean |
|---|---:|---:|---:|---:|---:|
| Windows group 0 | 43,238,352 | 1,176,109 | 42,526,315 | 452,177 | 783.65 → 813.41 |
| macOS | 43,216,905 | 1,137,347 | 42,295,447 | 767,432 | 824.08 → 735.54 |

The 25,000-row case still needs the final range index, so it retains its
approximately 6.8 MB worksheet/index footprint. The 1,000-row indexed control
also has no comparable memory benefit. The memory saving at 250,000 rows is
therefore a distinct boundary result, with the small and indexed costs retained.

A read-only review finds no actionable correctness regression. Its provisional
ZIP-length hypothesis is withdrawn after the [frozen Before/After reproduction](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-eligibility-review-reproduction.json) shows both builds reject
a 42,024,113-byte worksheet whose central-directory length is changed to
16,384 bytes, with the same truncated-XML error at position 16,385. Inspection
of the [.NET 8 inflater](https://github.com/dotnet/runtime/blob/v8.0.0/src/libraries/System.IO.Compression/src/System/IO/Compression/DeflateZLib/Inflater.cs#L77-L101)
and [.NET 10 inflater](https://github.com/dotnet/runtime/blob/v10.0.0/src/libraries/System.IO.Compression/src/System/IO/Compression/DeflateZLib/Inflater.cs#L77-L101) confirms that they limit decompressed
output to the declared length. This negative reproduction does not justify a
new stream wrapper.

The [Windows rotated packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-eligibility-rotated-windows.json),
[macOS rotated packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-eligibility-rotated-macos.json),
[Windows native packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-eligibility-native-windows.json),
[macOS native packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-eligibility-native-macos.json),
[Windows fresh-worker packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-eligibility-cold-windows.json)
and [macOS fresh-worker packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-eligibility-cold-macos.json)
retain source and binary fingerprints, input qualification and raw observations.
The first Windows rotated run failed its source-provenance guard and is excluded;
the qualified replacement runs use the frozen candidate checkout.

## Native index eligibility refinement — 2026-10-05

The V4 refinement limits the native dimension probe to known worksheet parts
larger than 8 MiB and no larger than the existing 64 MiB index-buffer limit.
Smaller parts use the original buffer path; known oversized parts decline
without inflation. Unknown lengths retain the header probe. This private
eligibility change preserves complete XML qualification and adds no public
option or dependency. It remains isolated from the integrated reader.

Nine controls cover dimensionless sheets with 1,000, 25,000, 45,000, 55,000,
250,000 and 1,000,000 rows, plus dimensioned indexable sheets with 1,000,
25,000 and 100,000 rows. The 45,000- and 55,000-row parts contain 7,310,188
and 8,951,348 decompressed bytes respectively, straddling the probe boundary.
Both sides use the same freshly built harness and dependencies; only the
Excel assembly and its symbols differ. Setup validates every projected field,
header, row count and the decoded worksheet, style and shared-string parts.

Native .NET 8 and .NET 10 jobs retain twelve measurements per observation on
Windows group 0 and macOS. Small and oversized native controls no longer show
the V3 probe's structural allocation cost. The probed controls still add about
14–15 KB per warmed operation. Timing remains mixed, including the following
unfavorable observations:

| Host and runtime | Case | V4/Before median ratio |
|---|---|---:|
| Windows group 0, .NET 10 | Indexed 100,000 | 1.507 |
| Windows group 0, .NET 10 | Dimensionless 55,000 | 1.337 |
| Windows group 0, .NET 10 | Dimensionless 250,000 | 2.056 |
| Windows group 0, .NET 8 | Dimensionless 45,000 | 1.702 |

Rotated .NET 10 comparisons retain twelve samples per side on each Windows
processor group and macOS. A separate control compares the identical baseline
assembly on both sides for the 25,000- and 250,000-row dimensionless cases and
the 100,000-row indexed case. Identical-build median ratios span 0.896–1.045
on Windows group 0, 0.942–1.117 on group 1, and 0.992–1.025 on macOS. These
controls expose timing variation; they do not remove candidate negatives.
The actual V4 comparison retains several ratios above 1.05 on Windows group 0,
including 1.267 at 25,000 rows and 1.163 at 250,000 rows. Group 1 has no ratio
above 1.05. On macOS, the indexed 1,000-row median changes from 7.396 to
10.571 ms, a ratio of 1.429; the other eight cases stay within 1.05.

Six fresh workers per side and case repeat the first-workbook memory
measurement, with fixture generation and full-field validation in another
process. The 250,000-row memory saving remains reproducible:

| Host | Before allocation bytes, mean | V4 allocation bytes, mean | Before held managed bytes, mean | V4 held managed bytes, mean | Before → V4 open/first-row ms, mean |
|---|---:|---:|---:|---:|---:|
| Windows group 0 | 43,239,661 | 1,172,339 | 42,527,707 | 449,560 | 772.96 → 756.99 |
| macOS | 43,229,472 | 1,139,344 | 42,313,480 | 774,560 | 922.49 → 820.66 |

The smaller and indexed controls retain the memory they need, and first-row
latencies remain mixed. These workers include JIT and engine initialization,
use a warm operating-system file cache and measure calling-thread allocation
separately from the sampler. Sampled peaks are lower bounds. Windows reuses
two process identifiers after their earlier workers exit; all 108 observations
come from separately launched and awaited worker processes, not reused workers.

Focused correctness passes 927 tests on each modern runtime on both hosts,
922 on Windows .NET Framework, and the owning library's .NET Standard build.
The [qualification and source packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-refinement-qualification.json)
and [source delta](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-refinement.patch)
identify the frozen candidate. The [Windows native packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-refinement-native-windows.json),
[macOS native packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-refinement-native-macos.json),
[Windows rotated and identical-build packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-refinement-rotated-windows.json),
[macOS rotated and identical-build packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-refinement-rotated-macos.json),
[Windows fresh-worker packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-refinement-cold-windows.json)
and [macOS fresh-worker packet](excel-csv-broad-throughput-2026-10-04/xlsx-native-index-refinement-cold-macos.json)
retain raw observations and fingerprints. The completed comparison does not
qualify this refinement as a portable default throughput improvement. Its
memory benefit remains a measured boundary result, and integration is held.

## DataTable normalization shortcut qualification — 2026-10-05

The D13 comparison rebuilds both sides from common integrated source. Before
restores the unconditional `null`-to-`DBNull` normalization loop; After skips
that loop when the row parser reports a complete row. The earlier input-buffer
reuse remains on both sides. Source inventories identify exactly one differing
C# file, `ExcelSheetReader.Range.DataTableRows.cs`; both snapshots share the
fresh harness and dependencies, with only the Excel assembly and symbols
changed. An older D12 compiled row-source checksum is not reconciled with its
recorded normalized source, so that compiled snapshot is not reused as Before.
This provenance limitation does not establish corruption or a source defect.

Fourteen cases cover 25,000-row numeric and sales tables, inferred decimal
columns, dense and sparse 8- and 65-column rows, reverse order, 100-row prefixes,
typed and range XML streaming, and a DataReader control outside the changed
row loop. Setup validates every value, its type, schema and row count. Native
.NET 8 and .NET 10 jobs on Windows group 0 and macOS retain 112 observations
and 1,344 measurements, with 24 warmups, twelve measurements, four invocations
per iteration, an unroll factor of one and no outlier removal.

Native timing is mixed, including large changes in controls that do not use
the modified loop. A .NET 10 PowerForge comparison rotates the matched sides
within each run on both Windows processor groups and macOS. It uses 24
warmups and 24 measurements with four workbook reads per sample. Identical
baseline assemblies on both sides provide four separate control cases. These
rotated suites retain another 108 observations and 2,592 measurements; retained
elapsed values describe the four-read batch, while the table below reports
matched median ratios.

| Case | Windows group 0 After/Before | Windows group 1 After/Before | macOS After/Before |
|---|---:|---:|---:|
| Numeric double DataTable, 25,000 rows | 0.957 | 0.968 | 1.069 |
| Inferred decimal DataTable, 25,000 rows | 1.031 | 0.992 | 1.094 |
| Sales DataTable, 25,000 rows | 1.052 | 1.028 | 0.989 |
| Dense 65-column rows, 1,000 rows | 1.151 | 1.145 | 0.992 |
| DataReader control, 25,000 rows | 1.013 | 1.021 | 1.041 |

The identical-build controls also vary. Their median ratios span 0.935–0.993
on Windows group 0, 0.891–1.000 on group 1, and 0.877–1.044 on macOS. Neither
those controls nor selected favorable cases establish a portable speed benefit
for the shortcut. Native allocation differences are small and inconsistent,
with no structural allocation saving. The unconditional normalization loop is
restored; the earlier buffer-reuse implementation remains.

The [Windows native packet](excel-csv-broad-throughput-2026-10-04/xlsx-datatable-normalization-native-windows.json),
[macOS native packet](excel-csv-broad-throughput-2026-10-04/xlsx-datatable-normalization-native-macos.json),
[Windows rotated and identical-build packet](excel-csv-broad-throughput-2026-10-04/xlsx-datatable-normalization-rotated-windows.json)
and [macOS rotated and identical-build packet](excel-csv-broad-throughput-2026-10-04/xlsx-datatable-normalization-rotated-macos.json)
retain the complete reports, raw measurements, common-source inventories,
binary fingerprints and output qualification. Unfavorable observations remain
in these packets. The removed shortcut is not counted as an accepted speed win.

The restored source passes the complete Excel suite on .NET 8 and .NET 10 on
both hosts: 5,614 passed and five existing skips per run, with no failures.
Windows also passes 1,038 focused reader and typed-mapping tests on .NET
Framework. The [restoration qualification packet](excel-csv-broad-throughput-2026-10-04/xlsx-datatable-normalization-qualification.json)
retains test-result fingerprints and the normalized row-source checksum,
matching the fresh Before source exactly. Previously integrated typed-presence,
formula fallback and cancellation fixes remain in this validated source.

## Typed XML staging screen and first-call memory rejection

A materialized typed-reader experiment stages the requested raw cells in one
pooled buffer, then maps the final header and merged logical rows. It avoids the
row-order qualification pass without replaying property setters when a later
row fragment replaces a value. The experiment changes three C# paths against
the conservative integrated source. It is retained on separate local branches;
it is not integrated or counted as an accepted performance improvement.

The eight-case screen covers narrow and wide ordered worksheets, successful
UTF-8 indexing controls, and fully mapped numeric rows. Actual .NET 8 and .NET
10 runs on Windows and macOS retain 64 observations and 768 measurements:
24 warmups, 12 measured iterations, four reads per iteration, unroll factor one,
and no outlier removal. Windows uses affinity mask 65535; macOS uses operating
system scheduling. Both sides use the same freshly built harness and
dependencies, differing only in the Excel assembly and its symbols. Setup
validates every projected value, row count and complete observation.

| Case | Windows .NET 10 After/Before median | Windows .NET 8 | macOS .NET 10 | macOS .NET 8 |
|---|---:|---:|---:|---:|
| Wide UTF-16 projection, Automatic, 5,000 rows | 0.651 | 0.614 | 0.564 | 0.551 |
| Wide UTF-16 projection, Sequential, 5,000 rows | 0.490 | 0.687 | 0.893 | 0.611 |
| Narrow UTF-16 projection, Automatic, 5,000 rows | 1.466 | 1.165 | 0.796 | 1.825 |
| Narrow UTF-16 projection, Sequential, 5,000 rows | 2.630 | 1.317 | 0.406 | 1.474 |
| Numeric objects, decimal mode disabled, 25,000 rows | 0.757 | 0.726 | 0.662 | 0.729 |
| Numeric objects, decimal mode enabled, 25,000 rows | 0.650 | 0.692 | 0.725 | 0.964 |

The wide cases improve on all four host/runtime combinations, but narrow-case
regressions and variation in unchanged UTF-8 controls prevent a portable speed
conclusion. macOS .NET 8 also reports higher allocation for several changed
cases. The [complete native summary](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-native-summary.json)
and [Windows](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-native-windows.json)
and [macOS](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-native-macos.json)
packets retain every case, including unfavorable observations and raw samples.

A separate .NET 10 first-call probe uses 96 fresh worker processes: three
repetitions of each of the eight cases on each host and each source version.
Before/After order alternates. A producer independently validates every
projected field and verifies identical decoded worksheet, style and
shared-string parts. Each worker performs its first complete typed read against
that frozen input, validates the complete returned observation, then measures
memory after the benchmark releases the reader and materialized result.

| Numeric objects, decimal mode disabled, 25,000 rows | Windows Before | Windows After | macOS Before | macOS After |
|---|---:|---:|---:|---:|
| Mean calling-thread allocation, bytes | 3,634,029 | 8,855,944 | 3,619,715 | 8,839,395 |
| Mean sampled managed-memory increase, bytes | 3,878,160 | 8,614,885 | 3,974,589 | 8,610,016 |
| Mean managed-memory increase after return and GC, bytes | 291,979 | 293,411 | 564,211 | 5,791,536 |

The first call adds about 5.2 MB of allocation on both hosts and more than
doubles the sampled managed-memory increase in these numeric controls. Pool
retention differs between hosts; the after-return figures are observations of
these processes, not a universal retention guarantee. The probe includes JIT,
type initialization and reflection invocation. Calling-thread allocation
excludes the sampler, sampled peaks are lower bounds, and the producer warms
operating-system file caches. These figures do not describe warmed throughput.

The [Windows memory packet](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-memory-windows.json)
and [macOS memory packet](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-memory-macos.json)
retain all 96 observations, fixture qualification, source/binary inventories and
runner fingerprints. The [source patch](excel-csv-broad-throughput-2026-10-04/typed-xml-staging.patch)
and [disposition](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-disposition.json)
record the rejected candidate. The preserved memory [producer](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-memory-fixtures.ps1),
[worker](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-memory-worker.ps1)
and [suite](excel-csv-broad-throughput-2026-10-04/typed-xml-staging-memory-suite.ps1)
show the measurement boundary.

Focused correctness passes on both modern runtimes and both hosts contain
1,043 passed tests per run; Windows .NET Framework contains 1,038. The candidate
is rejected before full-suite or integration review qualification because its
buffer adds a demonstrated memory cost and its narrow performance is mixed.
The conservative integrated typed-reader implementation remains the baseline.

## Worksheet encoding eligibility and first-call memory

The worksheet buffer owner checks the first 256 bytes against the existing
UTF-8 indexer's eligibility rule before renting a worksheet-sized buffer. A
declined UTF-16 worksheet retains the XML reader path. Accepted worksheets
resume the same ZIP stream and retain exact declared-length and trailing-byte
validation. The same rule applies before background worksheet prefetch rents
its full buffer. This changes no public option, encoding support or dependency.

The complete native matrix contains 76 cases: all 64 typed combinations of
1,000/5,000 rows, four/65 physical columns, ordered/reversed rows, UTF-8/UTF-16
and Automatic/Sequential/Parallel/streaming reads; two fully mapped numeric
object controls; two public-reader controls; and eight numeric public-reader
encoding/prefetch combinations. Actual .NET 8 and .NET 10 on Windows and macOS
retain 608 observations and 7,296 measurements. The policy is 24 warmups,
12 retained iterations, four invocations, unroll factor one and no outlier
removal. Every setup validates values, types, schema, row count and the complete
returned observation. Both sides use identical harness and dependency bytes;
five Excel C# paths differ, including one new encoding helper.

The [Windows native packet](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-native-windows.json),
[macOS native packet](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-native-macos.json)
and [complete native summary](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-native-summary.json)
retain every favorable and unfavorable observation. Timing varies materially:
the wide ordered UTF-16 Automatic case has After/Before medians of 1.115 on
Windows .NET 10, 1.850 on Windows .NET 8, 0.879 on macOS .NET 10 and 0.959 on
macOS .NET 8. These runs do not establish a portable speed improvement.

A second lane rotates Before/After order across 28 cases and runs eight
identical-build controls on two Windows affinity placements and macOS. It
retains 216 observations and 5,184 samples, with 24 warmups, 24 retained samples
and four complete reads per sample. Both hosts use the same fingerprinted
PowerForge controller payload. The identical-build median ratios span
0.888–1.089 on Windows mask 65535, 0.885–1.056 on mask 4294901760 and
0.983–1.147 on macOS. Candidate ratios span 0.860–1.107, 0.817–1.079 and
0.752–1.171 respectively. The [Windows](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-controls-windows.json),
[macOS](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-controls-macos.json)
and [comparison summary](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-controls-summary.json)
retain the complete reports and controller fingerprints.

Six cases with unresolved timing evidence receive a three-engine diagnostic:
two independently loaded copies of the baseline plus the candidate, rotated
within every iteration. Each host retains 18 observations with 48 samples per
engine/case, 24 warmups and four reads per sample. The public 25,000-row
candidate median ratios are 0.938 on Windows and 0.974 on macOS, compared with
identical-baseline control ratios of 0.979 and 0.968. The prior large
public-reader slowdown does not repeat in this diagnostic. The 1,000-row
candidate remains slower than Before by 5.95% and 4.16%, while the second
baseline differs by -5.00% and -8.27%. Small-reader timing remains mixed; these
controls neither erase the negative measurements nor prove a general speed
gain. The [Windows diagnostic](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-three-way-windows.json)
and [macOS diagnostic](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-three-way-macos.json)
retain all 1,728 samples.

The separate first-call memory lane covers 20 cases, each source version, both
modern runtimes and both hosts, with three fresh worker processes per
combination: 480 measured workers. Producers validate every selected field and
verify identical decoded worksheet/style/shared-string bytes. Actual .NET 8
also runs 80 separate complete-field validation workers before measurement.
Measurement workers do not run a benchmark read or setup before their first
complete operation. Before/After order alternates. All-thread allocation
includes the sampler and prefetch worker; caller-thread allocation is retained
separately.

| Wide ordered UTF-16 typed read, 5,000 rows, Automatic | All-thread allocation Before, MB | After, MB | Managed increase after return/GC Before, MB | After, MB |
|---|---:|---:|---:|---:|
| Windows .NET 10 | 34.944 | 1.360 | 33.882 | 0.291 |
| Windows .NET 8 | 35.396 | 1.800 | 34.242 | 0.653 |
| macOS .NET 10 | 35.376 | 1.763 | 34.233 | 0.562 |
| macOS .NET 8 | 35.637 | 2.038 | 34.531 | 0.872 |

| Numeric UTF-16 public read, 25,000 rows, prefetch enabled | All-thread allocation Before, MB | After, MB |
|---|---:|---:|
| Windows .NET 10 | 6.142 | 1.919 |
| Windows .NET 8 | 6.504 | 2.276 |
| macOS .NET 10 | 6.338 | 2.111 |
| macOS .NET 8 | 6.658 | 2.454 |

MB here means 1,000,000 bytes; values are means of three fresh workers per side.
The [memory summary](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-memory-summary.json)
retains every case, range, caller/all-thread allocation, sampled peak and
after-return figure. Raw packets are retained for [Windows .NET 10](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-memory-windows-net10.json),
[Windows .NET 8](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-memory-windows-net8.json),
[macOS .NET 10](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-memory-macos-net10.json)
and [macOS .NET 8](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-memory-macos-net8.json).
The fully mapped numeric controls retain their existing approximately 3.6 MB
caller allocation on .NET 10. The discarded XML-staging experiment's 5.2 MB
numeric first-call penalty is absent. Not every control improves: macOS
.NET 10's small UTF-8 numeric/public cases have approximately 67 KB/33 KB higher
mean caller allocation, with candidate ranges spanning the lower values seen
on the baseline. Those observations remain in the evidence.

The sampler polls every five milliseconds, so peak figures are observed lower
bounds. JIT/type initialization and reflection invocation are included; the
producer warms operating-system file caches. Readers and materialized results
are released before after-return/GC measurement. These values do not describe
live-reader retention, a worst-case memory budget or warmed throughput.

The change qualifies as a first-call allocation and retained-memory improvement
for ineligible buffered worksheets, including enabled prefetch. Timing remains
mixed across the representative matrix. Correctness passes contain 1,071
focused and 5,624 full-suite tests on each modern runtime and host, with five
existing full-suite skips and no failures. Windows .NET Framework passes 1,066
focused tests; the integrated source builds for .NET Standard 2.0. One full
read-only review and its targeted prefetch confirmation found no actionable
defects. The [qualification packet](excel-csv-broad-throughput-2026-10-04/worksheet-encoding-qualification.json)
records source, test and review boundaries. Large-reader throughput, earlier
typed-reader regressions, mixed/file CSV performance, Linux and XLSX write gaps
remain part of the broader performance work.

## Current public typed Excel reads

The refreshed .NET 10 comparison consumes every field of every row through the
public APIs. It covers 1,000, 25,000, 250,000 and 1,000,000 data rows with integer,
decimal, date and Boolean columns. Setup checks every header and typed value;
decoded worksheet, style and shared-string bytes match across engines. Each of
the three engines has 16 warmups and 24 retained samples in rotated order, with
one complete read per sample and no outlier removal. Windows uses two fixed CPU
placements; macOS uses operating-system scheduling. There are 36 observations
and 864 retained measurements.

The OfficeIMO source matches the locally integrated worksheet encoding change.
Comparison packages are Sylvan.Data.Excel 0.5.8 and ExcelReader.NET 5.1.1, with
their exact package license evidence retained. These are isolated benchmark
dependencies. This lane measures complete-read throughput, without an
allocation or first-row comparison claim.

| Million-row placement | OfficeIMO median | Sylvan median | ExcelReader.NET median | OfficeIMO / Sylvan | OfficeIMO / ExcelReader.NET |
| --- | ---: | ---: | ---: | ---: | ---: |
| Windows, mask 65535 | 2,031.3 ms | 831.0 ms | 243.8 ms | 2.44 | 8.33 |
| Windows, mask 4294901760 | 2,072.9 ms | 833.3 ms | 251.1 ms | 2.49 | 8.25 |
| macOS | 1,756.7 ms | 799.8 ms | 258.5 ms | 2.20 | 6.80 |

The large-reader throughput gap remains material on both hosts. Smaller scans
also remain slower in this lane. The [complete summary](excel-csv-broad-throughput-2026-10-04/public-reader-current-summary.json)
retains every size and placement; raw observations are available for
[Windows mask 65535](excel-csv-broad-throughput-2026-10-04/public-reader-current-windows-65535.json),
[Windows mask 4294901760](excel-csv-broad-throughput-2026-10-04/public-reader-current-windows-4294901760.json)
and [macOS](excel-csv-broad-throughput-2026-10-04/public-reader-current-macos.json).
The [runner](excel-csv-broad-throughput-2026-10-04/public-reader-current-runner.ps1),
[case definitions](excel-csv-broad-throughput-2026-10-04/public-reader-current-cases.json)
and [package license fingerprints](excel-csv-broad-throughput-2026-10-04/public-reader-current-licenses.json)
record the comparison boundary.

Profiles of five complete million-row reads show worksheet validation and value
access on the decompression path. Their source predates the encoding change;
the measured input is UTF-8. The [Windows profile summary](excel-csv-broad-throughput-2026-10-04/public-reader-profile-windows.json)
contains sampled allocation events, while the [macOS summary](excel-csv-broad-throughput-2026-10-04/public-reader-profile-macos.json)
contains thread-time samples without allocation events. These are profiler
observations, rather than exact allocation totals or a proven contention cost.

A separate 64 KiB stream-buffering experiment was rejected. Its .NET 10 screen
includes six cases on each host, 24 observations and 288 retained measurements.
The million-row After/Before median ratios are 1.003 on Windows and 1.028 on
macOS, while the 1,000-row control adds 65,736 allocated bytes per operation on
both hosts. It does not improve the principal target and adds a portable
allocation cost. The [disposition](excel-csv-broad-throughput-2026-10-04/xml-stream-buffering-screen-disposition.json),
[Windows observations](excel-csv-broad-throughput-2026-10-04/xml-stream-buffering-screen-windows.json),
[macOS observations](excel-csv-broad-throughput-2026-10-04/xml-stream-buffering-screen-macos.json)
and [rejected patch](excel-csv-broad-throughput-2026-10-04/xml-stream-buffering-screen.patch)
are retained. The change is removed from both experiment branches and is absent
from the integration branch.

## Mixed-data and file CSV writes

The current CSV matrix covers nullable mixed typed inputs at 25,000 and 100,000
rows, with ordinary, quoted and multiline fields. It compares sequential and
parallel OfficeIMO DataReader writes with Sylvan.Data.Csv. File writes cover
seven shapes in both quoting modes: short ASCII, short Unicode, dense JSON,
quote runs, long notes, typed values and 25,000-row mixed JSON. They compare with
CsvHelper using identical stream buffers, output bytes and complete field
validation. Mixed exports validate every decoded field; formatting can differ
when the decoded typed values are equal. File creation, close and flushing to
the operating system are included; durable media flushing is excluded.

Both actual .NET 8 and .NET 10 run on Windows, macOS and Linux under WSL2, with
two fixed Windows CPU placements. Linux runs Ubuntu 24.04.3 on the Windows host;
its results describe that WSL2 environment, with operating-system scheduling.
The captured 936 source/project files match across all three environments. Each
case has 24 warmups, 12 retained measurements, four invocations and no outlier
removal. The complete matrix contains 368 observations, 4,416 retained
measurements and 208 matched comparisons. Package versions are CsvHelper
33.1.0 and Sylvan.Data.Csv 1.4.4.

OfficeIMO has lower medians in 167 of 208 comparisons. File writes allocate less
in 108 of 112 comparisons; mixed-data writes allocate less in 48 of 96, with small
extra allocations in the remaining mixed cases. Negative timings include short
Unicode file output, some long-note writes and selected mixed/quoted cases.
The measurements support workload-specific findings and retain the slower
cases; they do not establish an overall CSV speed or allocation lead.

Linux alone has lower OfficeIMO medians in 40 of 52 comparisons and lower
allocations in all 28 file-write comparisons. The .NET 8 quoted mixed-data
sequential write has a 1.877 median ratio, while the .NET 10 parallel counterpart
has a 1.522 ratio. These negative cases remain in the matrix.

The [complete three-environment summary](excel-csv-broad-throughput-2026-10-04/csv-mixed-file-current-linux-summary.json)
contains each timing/allocation pair and source fingerprints. Raw native
observations are retained for [Windows](excel-csv-broad-throughput-2026-10-04/csv-mixed-file-current-windows.json)
and [macOS](excel-csv-broad-throughput-2026-10-04/csv-mixed-file-current-macos.json),
with the additional [Linux WSL2 packet](excel-csv-broad-throughput-2026-10-04/csv-mixed-file-current-linux.json)
and the original [Windows/macOS summary](excel-csv-broad-throughput-2026-10-04/csv-mixed-file-current-summary.json).
The evidence includes the [runner](excel-csv-broad-throughput-2026-10-04/csv-mixed-file-current-runner.ps1)
and [summary generator](excel-csv-broad-throughput-2026-10-04/csv-mixed-file-current-summarize.ps1).
Further changes require comparable before/after proof across these workloads,
including the negative cases.

A separate experiment broadened completed-record batching to unformatted
multi-character delimiters. It is rejected for integration. Its native matrix
uses all 14 file cases on both modern runtimes and three placements: 168
observations and 2,016 retained measurements, with four complete writes per
sample. The two short-Unicode target cases improve on macOS .NET 10, but their
macOS .NET 8 median ratios are 1.146 and 2.068. Target allocations increase by
4,102 bytes per operation on Windows .NET 10 and macOS .NET 8.

The additional .NET 10 controls rotate Before, an identical second baseline,
and After within each iteration. Six cases on each host yield 36 observations
and 1,728 retained measurements. Windows target After/Before medians are 1.116
and 1.081; identical-baseline Control/Before ratios are 1.170 and 1.073. macOS
target ratios are 0.846 and 0.884, with controls at 0.820 and 0.985. This evidence
does not establish a portable gain. The production change is backed out on both
experiment branches and absent from the integration branch. The
[disposition and qualification](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching-disposition.json),
[native summary](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching-native-summary.json),
[rotated control summary](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching-control-summary.json)
and [rejected patch](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching.patch)
retain the candidate boundary and slower cases. Correctness and independent
review qualify the implementation, but do not establish a performance benefit.
Raw packets are retained for [Windows mask 65535](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching-native-windows-65535.json),
[Windows mask 4294901760](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching-native-windows-4294901760.json)
and [macOS](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching-native-macos.json),
with rotated controls for [Windows](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching-control-windows.json)
and [macOS](excel-csv-broad-throughput-2026-10-04/csv-text-delimiter-batching-control-macos.json).
The additional cancellation, partial-record and writer-reuse cases remain in
the correctness suite: all 28 writer regression cases pass against the frozen
qualified CSV baseline after the experiment is backed out.

## Undimensioned worksheet indexing

The indexed reader can infer a bounded proposal from the first populated row
and the final explicit row when the worksheet omits `dimension`. It qualifies
the complete XML, cells, grid coordinates and actual used bounds before exposing
rows. UTF-16, unsupported prefixes, later wider cells and proposals above the
existing cell or byte budgets retain their established paths. The private
optimization adds no public options or runtime dependencies.

The native matrix covers twelve cases on Windows and macOS, on actual .NET 8
and .NET 10: 96 observations and 1,152 retained measurements. It includes four
public sizes from 1,000 to 1,000,000 rows, UTF-8 and UTF-16 numeric data, prefetch,
and wide Automatic/Sequential reads. The additional eight-case controls rotate
Before, an identical second baseline, and After within each iteration. Four
host/runtime packets contain 96 observations and 2,304 retained measurements,
with 24 warmups and 24 retained samples per engine. Every setup validates all
selected fields and matching decoded worksheet, style and shared-string bytes.

| Public read, 25,000 rows | After/Before median | Identical-baseline Control/Before |
| --- | ---: | ---: |
| Windows .NET 10 | 0.547 | 0.998 |
| macOS .NET 10 | 0.538 | 1.065 |
| Windows .NET 8 | 0.471 | 0.968 |
| macOS .NET 8 | 0.540 | 1.001 |

These controls support a portable improvement for eligible medium worksheets.
They do not establish a large-sheet throughput improvement. Native negative
cases remain visible, including wide Sequential UTF-8 ratios of approximately
4.58 on Windows .NET 10 and 4.02 on macOS .NET 8. The rotated Windows .NET 10
wide Sequential ratio is 1.096, and macOS .NET 10 UTF-16 numeric reads have a
1.137 ratio. Large public scans have mixed ratios around their baselines.
The [native summary](excel-csv-broad-throughput-2026-10-04/dimensionless-index-native-summary.json)
and [rotated summary](excel-csv-broad-throughput-2026-10-04/dimensionless-index-control-summary.json)
retain every case and the corresponding raw packet fingerprints.

Fresh-process memory evidence contains 144 measured workers and 24 additional
.NET 8 complete-field validation workers. Three workers per side cover four
public sizes and the UTF-8/UTF-16 numeric controls. The 250,000-row case still
allocates approximately 42.8 MB on the calling thread on Windows .NET 10 and
retains approximately 42.3 MB after return/GC in both builds. Its 250,001 rows,
including the header, span 1,000,004 cells and exceed the existing index budget;
the full worksheet buffer is acquired before that proposal is declined.
This change leaves that allocation gap open. The
[memory summary](excel-csv-broad-throughput-2026-10-04/dimensionless-index-memory-summary.json)
retains caller/all-thread allocations, after-return figures and sampled peaks.
Small and medium controls have modest allocation reductions, with overlapping
ranges and some higher individual or mean figures retained. The five-millisecond
sampler provides peak lower bounds; this is first-complete-operation evidence,
including initialization and pool retention, rather than a warmed speed claim.

Native .NET 8 uses 8.0.31 on Windows and 8.0.23 on macOS. The actual .NET 8
PowerShell controls use 8.0.21 on both hosts; .NET 10 uses 10.0.12. Runtime patch
versions are recorded rather than treating these as identical environments.
Raw native, rotated and memory packets and their captured runner scripts are
retained beside the summaries. The qualification source is `1b0c3b146`; its
normalized Excel source is integrated at `b57455a0c`.

## Semicolon and tab quote-character searches

CSV writers use the existing vectorized character-search mechanism for
semicolon and tab separators on .NET 8 and later. Both string and span paths
preserve the earliest delimiter, quote, CR or LF position. The two search sets
initialize once on first use. Older targets retain their existing scalar scan.
This changes private implementation only and preserves quoting and output bytes.

The native matrix covers 38 cases on both modern runtimes, two Windows CPU
placements and macOS: 456 observations, 5,472 retained measurements and 228
Before/After pairs. Cases include seven complete-file shapes in both quoting
modes, long and short semicolon/tab text, and comma/multicharacter controls.
Every implementation produces equivalent text or bytes and validates every
decoded field. Warmed allocation is exactly equal in all 228 matched pairs.
This evidence does not measure the cold allocation of the search sets.

Ten-case rotated controls on both runtimes and hosts add 120 observations and
5,760 retained measurements. They use 24 warmups, 48 retained samples and four
complete writes per sample. Before and Control load identical baseline bytes.

| Long text, AsNeeded quoting | Windows .NET 10 After/Before | macOS .NET 10 | Windows .NET 8 | macOS .NET 8 |
| --- | ---: | ---: | ---: | ---: |
| Semicolon, plain | 0.459 | 0.531 | 0.509 | 0.548 |
| Semicolon, containing delimiter | 0.680 | 0.677 | 0.641 | 0.775 |
| Tab, plain | 0.447 | 0.546 | 0.487 | 0.517 |
| Tab, containing delimiter | 0.797 | 0.667 | 0.658 | 0.688 |

The target cases improve in all four rotated packets. The native matrix also
retains slower observations, including semicolon targets in selected runs and
short Unicode complete-file output. Forward rotated macOS Unicode ratios are
1.12 on .NET 10 and 1.13 on .NET 8. That file case uses a multicharacter separator
and does not reach the changed search paths. A three-case reversed-role check
loads candidate bytes for Before/Control and baseline bytes for After; its four
packets contain 36 observations and 1,728 retained measurements. Candidate over
baseline Unicode ratios become 1.016 on macOS .NET 10 and 0.929 on .NET 8.
Identical-copy ratios also vary. The role reversal does not reproduce a stable
Unicode slowdown; all slower observations remain in the evidence. It does not
prove that every unchanged workload has zero regression.

The [native summary](excel-csv-broad-throughput-2026-10-04/csv-delimiter-search-native-summary.json),
[rotated summary](excel-csv-broad-throughput-2026-10-04/csv-delimiter-search-control-summary.json)
and [reversed-role summary](excel-csv-broad-throughput-2026-10-04/csv-delimiter-search-reversed-summary.json)
retain all cases. Their raw packets preserve full source, assembly and measurement
provenance. One discarded forward Windows run failed the source-provenance guard
because it was launched from a changing checkout; it supplies no qualified
measurements. The stable rerun is the retained Windows .NET 10 packet.
The qualification source is `0bed96b2a`; its normalized CSV source is integrated
at `b57455a0c`. Linux qualification of this change remains open.

The combined source passes 5,629 Excel tests with five existing skips and 681
CSV tests on each modern runtime and host. Windows .NET Framework passes 938
reader-focused Excel tests and all 487 CSV tests. Both owners build for .NET
Standard 2.0 with zero warnings or errors. The
[integration qualification](excel-csv-broad-throughput-2026-10-04/index-and-delimiter-integration-qualification.json)
records source equality, exact test fingerprints and independent review scope.
The index review reproduced and closed an introduced grid-coordinate defect;
the delimiter-search review found no actionable defects. These are local source
and correctness findings. The remaining large-reader, memory, writer and CSV
read comparisons remain part of the broad performance investigation.
