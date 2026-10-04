# Large XLSX typed reads — 2026-10-04

Commit `cc694e167` avoids allocating a worksheet buffer when the package index
already declares that the worksheet exceeds the indexed reader's 64 MiB limit.
The reader uses its existing streaming fallback, preserving validation before
returning rows. The baseline is `9fb634005`, following the
[broader CSV/XLSX allocation changes](officeimo.excel-csv-broad-throughput-2026-10-04.md).

This removes about 128 MiB of managed allocation from opening the million-row
fixture. A subsequent change, `8dc37d02a`, discovers worksheet bounds during
complete reader validation and removes another 153 MiB from that operation.
Buffered numeric decoding subsequently removes another 77 MiB from a complete
million-row scan. The measurements below keep these stages separate. Large-sheet
allocation and opening costs remain, and the contended timing runs do not
establish a portable speedup.

## Workload and allocation

The generated workbooks contain 25,000, 250,000, or 1,000,000 rows with an integer,
decimal, date/time, and boolean. `ExcelDocument.WriteRows` writes explicit cell
references without shared strings. Its single-pass output omits the optional
worksheet dimension. The million-row file is 37,033,966 bytes and contains
170,561,620 bytes of worksheet XML.

Setup validates every header and typed value outside timing. Each measured
operation opens the workbook, reads all four fields of the requested rows,
validates the count and checksum, and disposes the reader. First-row cases retain
OfficeIMO's eager worksheet validation; they are not ranked against readers
with different validation policies.

| Operation | Rows | Before allocation | After allocation |
| --- | ---: | ---: | ---: |
| Open and scan all rows | 25,000 | 9.28 MiB | 9.28 MiB |
| Open and read first row | 25,000 | 9.28 MiB | 9.28 MiB |
| Open and scan all rows | 250,000 | 201.24 MiB | 201.24 MiB |
| Open and read first row | 250,000 | 96.26 MiB | 96.26 MiB |
| Open and scan all rows | 1,000,000 | 943.11 MiB | 815.17 MiB |
| Open and read first row | 1,000,000 | 516.17 MiB | 388.23 MiB |

The million-row reductions are 13.6% for the complete scan and 24.8% for the
first-row operation. The separate 65K fourteen-column corpus remains about
90 KiB in this snapshot driver. These are BenchmarkDotNet allocated bytes per
operation, not live retained objects or sampled process-memory peaks.

All fourteen before/after cases complete eight measured iterations after eight
warmups, with one operation per iteration and no outlier removal. The Windows
run uses .NET 10.0.12, Normal priority, and the previously verified `0xFFFF`
processor group. This mask is specific to the measured workstation. Other work
caused substantial timing variation, including million-row first-row samples
above ten seconds; the complete timing observations remain in the evidence.

## Consolidated worksheet scans

The second comparison uses `ee2930812` as the baseline and `8dc37d02a` as the
candidate. The validation scan discovers used-range bounds without constructing
another cell-reference string for range discovery. Bounds become available only
after the complete XML scan and its final cancellation check. Table-bearing
worksheets and unsupported XML layouts retain the existing discovery path.

The private coordinate accumulator is shared with ordinary used-range discovery.
Style, shared-string, shared-formula, and trailing-XML validation remains eager.
No public API or dependency changes are involved.

| Operation | Rows | Before allocation | After allocation |
| --- | ---: | ---: | ---: |
| Open and scan all rows | 25,000 | 9.28 MiB | 6.97 MiB |
| Open and read first row | 25,000 | 9.28 MiB | 5.74 MiB |
| Open and scan all rows | 250,000 | 201.26 MiB | 163.34 MiB |
| Open and read first row | 250,000 | 96.28 MiB | 58.36 MiB |
| Open and scan all rows | 1,000,000 | 815.17 MiB | 662.67 MiB |
| Open and read first row | 1,000,000 | 388.24 MiB | 235.75 MiB |
| Fourteen-column corpus | 65K | 10.03 MiB | 10.03 MiB |

All fourteen cases complete five warmups and five measured iterations with one
operation per iteration. Processor placement, priority, runtime, output
validation, and outlier retention match the first stage. The million-row
reduction is 18.7% for a complete scan and 39.3% for the first-row operation.
Relative to the earlier pre-buffer-guard observations, these operations allocate
about 30% and 54% less, respectively.

The 65K control allocates equally in both versions but differs substantially
from the earlier 90 KiB observation. A follow-up uses 24 warmups, twelve measured
iterations, and four operations per iteration. The 65K control then allocates
91,486 bytes per operation in both versions. Both 25,000-row operations measure
about 9.28 → 5.73 MiB, a 38.2% reduction. The initial and follow-up observations
demonstrate that the short warmup is insufficient for a stable allocation budget
on these smaller cases; they do not establish the cause of the variation.

All six follow-up cases pass. Timing remains inconsistent: the candidate's
25,000-row first-row mean is slower in the follow-up despite lower allocation.
Both the [initial observations](excel-large-typed-read-2026-10-04/range-scan-native.json)
and [longer-warmup observations](excel-large-typed-read-2026-10-04/pool-check-native.json)
remain available. No throughput improvement is claimed from these contended runs.

## Buffered numeric XML decoding

The next comparison uses `6d36a99c4` as its source baseline, including the scan
consolidation above. Candidate `d16c7d2ce` parses cached numeric XML values into the
existing reusable character buffer before converting the stored double to a
decimal. It preserves date handling, numeric precision, overflow fallback, and
the caller's formula-result policy. The
[numeric provenance](excel-large-typed-read-2026-10-04/numeric-provenance.json)
identifies the candidate source and measured assemblies.

The added numeric lane reads two columns and 25,000 rows through DataReader,
range, and DataTable APIs, each with double and decimal options. UTF-16 input
exercises the XML fallback independently of the indexed reader's size boundary.
Setup validates every header, value, type, and row count. These are separate
before/after contracts; the different result shapes are not ranked against one
another.

| Numeric API and mode | Before allocation | After allocation |
| --- | ---: | ---: |
| DataReader, decimal | 8.55 MiB | 6.85 MiB |
| Range, decimal | 6.35 MiB | 4.64 MiB |
| DataTable, decimal | 11.01 MiB | 9.29 MiB |
| DataReader, double | 6.47 MiB | 6.45 MiB |
| Range, double | 4.25 MiB | 4.25 MiB |
| DataTable, double | 8.91 MiB | 8.91 MiB |

These cases and five adjacent controls complete 24 warmups, twelve measured
iterations, and four operations per iteration. The ordinary 25K typed scan and
first-row operation, 65K corpus, and 2,500-row object/DataTable reads allocate
essentially the same amount before and after. The decimal reductions are 19.9%,
27.0%, and 15.6%, respectively.

The larger typed workload completes five warmups and five measured iterations
with one operation per iteration:

| Operation | Rows | Before allocation | After allocation |
| --- | ---: | ---: | ---: |
| Complete scan | 250,000 | 163.33 MiB | 145.02 MiB |
| Complete scan | 1,000,000 | 662.67 MiB | 585.62 MiB |
| Open and read first row | 250,000 | 58.35 MiB | 58.35 MiB |
| Open and read first row | 1,000,000 | 235.73 MiB | 235.73 MiB |

The complete-scan reductions are 11.2% and 11.6%. First-row allocation is
unchanged because that operation decodes few numeric values. Across all three
stages, the million-row complete scan allocates about 38% less than the original
943.11 MiB observation. This does not establish a reduction in peak or retained
memory.

All [30 native cases](excel-large-typed-read-2026-10-04/numeric-native.json)
complete. Timing remains mixed. A follow-up rotates all six numeric cases and
the 65K control across both verified processor groups, retaining
[672 measured samples](excel-large-typed-read-2026-10-04/numeric-rotated.json).
For example, decimal range reads have candidate/baseline mean ratios of 1.12
and 0.95 on the two groups. These observations do not establish a portable
speedup or a consistent throughput regression; unfavorable observations remain
in the packet. All runs use Normal priority and retain outliers.

## Coordinate buffering: allocation and throughput tradeoff

The initial coordinate-buffering checkpoint, `6a533b308`, reads ordinary row and
cell references into a reusable character buffer. It keeps long-reference and
diagnostic text intact and applies the same parsing helper across XML-backed
reader APIs. The baseline is `4c73bdcf3`; the
[initial provenance](excel-large-typed-read-2026-10-04/coordinate-initial-provenance.json)
pins both product assemblies and the identical benchmark harness.

| Operation | Rows | Before allocation | Initial candidate allocation |
| --- | ---: | ---: | ---: |
| Decimal DataReader, two numeric columns | 25,000 | 6.83 MiB | 1.81 MiB |
| Double DataReader, two numeric columns | 25,000 | 6.45 MiB | 1.43 MiB |
| Decimal range, two numeric columns | 25,000 | 4.64 MiB | 2.13 MiB |
| Double range, two numeric columns | 25,000 | 4.25 MiB | 1.74 MiB |
| Decimal DataTable, two numeric columns | 25,000 | 9.29 MiB | 6.78 MiB |
| Double DataTable, two numeric columns | 25,000 | 8.91 MiB | 6.40 MiB |
| Materialized typed objects | 2,500 | 3.02 MiB | 2.33 MiB |
| Materialized DataTable | 2,500 | 2.52 MiB | 1.84 MiB |
| Complete typed scan | 250,000 | 145.02 MiB | 52.16 MiB |
| Complete typed scan | 1,000,000 | 585.62 MiB | 207.94 MiB |
| Open and read first row | 250,000 | 58.35 MiB | 11.92 MiB |
| Open and read first row | 1,000,000 | 235.73 MiB | 46.87 MiB |

All [38 native cases](excel-large-typed-read-2026-10-04/coordinate-initial-native.json)
complete. Ordinary and prefixed shared-string controls have effectively unchanged
allocation. The 65K corpus again varies with pool state: the baseline reports
2.57 MiB in this run versus about 90 KiB in earlier longer-warmup observations.
That difference is not credited to coordinate buffering.

The allocation improvement does **not** qualify the initial helper as a
throughput improvement. The
[15-scenario rotated comparison](excel-large-typed-read-2026-10-04/coordinate-initial-rotated.json)
retains 1,440 samples, including slower materialized and numeric reads. A separate
[large-sheet rotation](excel-large-typed-read-2026-10-04/coordinate-initial-large-rotated.json)
retains 192 samples. On the two processor groups, candidate/baseline mean ratios
are 1.08/1.08 for million-row first-row opening and 1.07/1.09 for 250K first-row
opening. Complete scans have ratios of 1.04/1.07 at one million rows and
1.02/1.16 at 250K rows. These repeated slower observations require throughput
remediation; they are not discarded as host noise.

The full Windows .NET 10 non-performance suite passes 5,246 cases with five
skips. All nine coordinate-specific contract cases pass after adding three
diagnostic cases. They cover UTF-16, prefixed XML, entity-escaped and long
references, missing and empty references, and diagnostics after reading cell
content. An independent read-only review reports no actionable correctness or
API findings. Other runtime qualification applies to a stable revised candidate,
not to this initial timing experiment.

## Typed values and revised XML traversal

Candidate `b854ad35b` retains numeric and date serials in the current-row
cache until an object result is requested. Typed decimal and date getters avoid
boxing, and a later object getter still returns the configured canonical type.
The candidate also shares immutable XML schema names and avoids repeating
worksheet-structure checks for every validated cell. It includes the coordinate
buffering above; the comparison baseline remains `4c73bdcf3`, so the figures
measure their combined effect rather than attributing all savings to typed getters.

The added `TypedDataReader` lane reads the same two numeric fields through
`GetInt32` and `GetDecimal`. The existing `DataReader` lane uses `GetValue` and
remains an object-materialization control. Setup verifies complete values and
canonical object types for both double and decimal options. Small 2,500-row
cases exercise buffered readers; 25,000-row UTF-16 cases exercise streaming XML.

| Operation | Rows | Baseline allocation | Revised candidate allocation |
| --- | ---: | ---: | ---: |
| Typed numeric DataReader, decimal | 2,500 | 1.09 MiB | 0.55 MiB |
| Typed numeric DataReader, double | 2,500 | 1.05 MiB | 0.52 MiB |
| Typed numeric DataReader, decimal | 25,000 | 6.83 MiB | 0.28 MiB |
| Typed numeric DataReader, double | 25,000 | 5.31 MiB | 0.28 MiB |
| Object numeric DataReader, decimal | 25,000 | 6.83 MiB | 1.81 MiB |
| Object numeric DataReader, double | 25,000 | 6.45 MiB | 1.43 MiB |
| Complete four-column typed scan | 250,000 | 145.02 MiB | 23.54 MiB |
| Complete four-column typed scan | 1,000,000 | 585.62 MiB | 93.45 MiB |
| Open and read first row | 250,000 | 58.35 MiB | 11.91 MiB |
| Open and read first row | 1,000,000 | 235.73 MiB | 46.87 MiB |

The 25K decimal typed lane allocates 95.8% less; the complete million-row typed
scan allocates 84.0% less. Shared-string controls remain effectively unchanged.
The 65K corpus allocates about 90 KiB in both snapshots. These measurements do
not establish lower retained or peak memory.

All [46 native cases](excel-large-typed-read-2026-10-04/typed-values-native.json)
and 2,016 output-validated rotated samples complete. The smaller workloads use
24 native warmups, twelve iterations, and four operations per iteration; the
large workloads use five warmups, five iterations, and one operation. PowerForge
rotates both versions on each processor group with the settings recorded in the
[provenance](excel-large-typed-read-2026-10-04/typed-values-provenance.json).
Every sample and outlier remains available in the
[small](excel-large-typed-read-2026-10-04/typed-values-small-rotated.json) and
[large](excel-large-typed-read-2026-10-04/typed-values-large-rotated.json) packets.

Timing does not qualify this candidate as a general speedup. Candidate/baseline
mean ratios from the two processor groups are:

| Operation | `0xFFFF` | `0xFFFF0000` |
| --- | ---: | ---: |
| Million-row complete scan | 0.98 | 1.06 |
| Million-row first-row opening | 1.00 | 1.10 |
| 250K complete scan | 1.01 | 1.04 |
| 250K first-row opening | 0.91 | 1.04 |
| 25K typed numeric DataReader, double | 1.08 | 1.13 |
| 25K object numeric DataReader, double | 1.17 | 1.05 |
| 65K corpus | 1.11 | 1.03 |

Several workloads change direction between groups, and means and medians can
differ considerably. For example, the 2,500-row double typed lane has mean
ratios of 1.48 and 0.80, with median ratios of 1.13 and 0.94. Million-row
first-row medians are 0.98 and 1.06. Repeated slower observations in the numeric
and opening paths remain required throughput work; they are not erased by the
allocation savings or by faster native observations.

The preceding helper-inlining experiment was rejected. Its
[measurements](excel-large-typed-read-2026-10-04/coordinate-rejected-inline.json)
and [patch](excel-large-typed-read-2026-10-04/coordinate-rejected-inline.patch)
remain for reproduction. Separate
[validation traversal](excel-large-typed-read-2026-10-04/coordinate-validation-screen.json)
and [schema-name](excel-large-typed-read-2026-10-04/coordinate-schema-screen.json)
screens also retain slower observations; neither is represented as an
independently qualified speed improvement.

Getter-order validation also repairs pre-existing inconsistencies: typed reads
could change subsequent decimal object results, numeric access could discard
date interpretation, and buffered readers eagerly converted out-of-range date
serials before callers could retrieve their numeric value. Eight regression
cases fail before and pass after the change, covering UTF-8/UTF-16, the 1900/1904
date systems, and buffered/streaming paths. Independent review identifies a
stale primitive-cache check when an unsorted reader switches to buffered rows;
five of eight transition cases reproduce it, all eight pass after correction,
and the targeted review confirms the fix.

The final Windows .NET 10 non-performance suite passes 5,265 tests with five
skips. Focused reader checks pass 645 cases on .NET 8 and 640 on .NET Framework
4.7.2. A fresh Ubuntu 24.04/WSL .NET 10 build passes the same 645 focused cases.
The product builds for `netstandard2.0` and the benchmark builds for .NET 8 and
.NET 10 with zero warnings or errors. No public API or production dependency
changes are introduced. The sparse implicit-row position defect remains open
at this checkpoint.


## Implicit worksheet coordinates

Commit `15b5f2848` corrects row and column inference across twelve reader APIs,
including typed and buffered readers, ranges, rows, columns, dictionaries,
objects, DataTables, and cell enumeration. Missing row indices use the first
valid cell reference or the next sequential row. DOM readers infer omitted
cell references while traversing the existing row. Reading an attached live
worksheet preserves its model. Live dimensions, header lookup, and header-cache
validation use the same coordinate rules, including out-of-order rows.

The original 36 implicit-row fixtures fail before the change. Subsequent
fixtures reproduce nine missing-cell failures, an empty live header map, and
two unsorted-header failures. The final 96 coordinate cases pass on Windows
.NET 10, .NET 8, .NET Framework 4.7.2, and Ubuntu/WSL .NET 10. The full Windows
.NET 10 suite passes 5,359 tests with five skips before the final narrow
unsorted-header correction; 741 focused reader checks pass after that
correction. The broader pre-correction focused runs pass 739 tests on .NET 8
and Linux, and 734 on .NET Framework. Product `netstandard2.0` and benchmark
.NET 8/.NET 10 builds have no warnings or errors.

One independent read-only review and one targeted confirmation identify the
missing-column and header issues. All findings are reproduced and corrected.
The final unsorted-header correction has direct regression, cache-refresh, and
runtime proof; it does not receive an additional independent review.

This is a correctness checkpoint with a measured throughput cost. The XML
fallback makes a separate scan when it first encounters an omitted row index,
retaining only departures from sequential numbering. The six-workload
[case file](excel-large-typed-read-2026-10-04/implicit-coordinate-cases.json)
compares explicit coordinates, omitted row indices, and omitted row and cell
indices at 25,000 rows. Both versions validate every numeric value. The
[native packet](excel-large-typed-read-2026-10-04/implicit-coordinates-native.json)
contains twelve cases, and the
[rotated packet](excel-large-typed-read-2026-10-04/implicit-coordinates-rotated.json)
contains 576 successful samples across both processor groups.

| Coordinate layout and API | Mean ratio, `0xFFFF` | Mean ratio, `0xFFFF0000` |
| --- | ---: | ---: |
| Explicit, range | 1.10 | 1.00 |
| Explicit, typed reader | 1.01 | 1.02 |
| Omitted rows, range | 1.61 | 1.62 |
| Omitted rows, typed reader | 1.20 | 1.22 |
| Omitted rows and cells, range | 1.65 | 1.50 |
| Omitted rows and cells, typed reader | 1.43 | 1.28 |

Ratios above one are slower. The implicit-coordinate regressions repeat on
both groups and remain required performance work. Native managed allocation
increases by about 16–21 KiB in those lanes; explicit-coordinate allocation
remains effectively unchanged. The comparison uses the preceding typed-reader
commit `b854ad35b` as baseline. The
[provenance](excel-large-typed-read-2026-10-04/implicit-coordinates-provenance.json)
records both binaries, benchmark hash, test reports, and review boundaries.

A fresh five-scan million-row
[profile](excel-large-typed-read-2026-10-04/typed-values-current-profile.json)
attributes 488.8 million of 489.9 million sampled allocated bytes to strings.
Validation and current-value reading account for approximately equal shares.
These are weighted diagnostic samples, not exact allocation accounting or CPU
time. They identify XML attribute/value handling as the next allocation
investigation alongside reuse of existing coordinate scans.

## Remaining allocation owners

A profile after the buffer guard, before scan consolidation, attributes roughly
190 MiB per operation to used-range discovery and 198 MiB to worksheet validation
across three million-row scans. These
are weighted allocation samples, not exact accounting. Both scans inspect cell
references independently. Typed value decoding and row traversal account for
most of the remaining allocation. The profile's thread-stack counts must not be
interpreted as precise CPU time.

The [spreadsheet roadmap](../ROADMAP.md#spreadsheet-and-csv-delivery-order)
retains large-sheet decoding, general read/export throughput, and portable timing
and memory qualification. Validation also exposes a pre-existing disagreement
between used-range inference and reader row positions when a row omits its index
and its cells reference a later row. The saved baseline reproduces it; the scan
consolidation does not fix or worsen that separate defect.

## Validation and reproduction

Focused DataReader tests pass on Windows: 250 on .NET 10, 250 on .NET 8, and 245
on .NET Framework 4.7.2. The benchmark project builds for .NET 8 and .NET 10;
the product builds for `netstandard2.0`, with zero warnings or errors.
Before/after preflight compares complete generated row sets at every size and
the hash-pinned 65K corpus. The earlier Linux/WSL results cover the broader
implementation, not a new Linux run of this buffer decision.

For scan consolidation, the complete Windows .NET 10 non-performance suite
passes 5,233 tests with five skips. Focused DataReader, used-range, and range-read
tests pass 323 cases on .NET 8 and 318 on .NET Framework 4.7.2; the
`netstandard2.0` product build has zero warnings or errors. The new fixtures
cover sparse and sequential implicit coordinates, ordinary and prefixed XML,
invalid metadata, and malformed trailing XML. One independent read-only review
of this change reports no actionable findings. The
[second-stage provenance](excel-large-typed-read-2026-10-04/range-scan-provenance.json)
records source/binary hashes, test outcomes, and the review fingerprint.
A fresh .NET 10 build under Ubuntu 24.04/WSL passes the same 323 focused tests.
That qualifies correctness on this Linux environment, not native-Linux timing.

For numeric decoding, the full Windows .NET 10 suite passes 5,240 tests with five
skips. Focused reader, range, and shared-string checks pass 608 tests on .NET 8
and 603 on .NET Framework 4.7.2. The product builds for `netstandard2.0` and the
benchmark builds for .NET 8 and .NET 10 with zero warnings or errors. A fresh
Ubuntu 24.04/WSL build passes the same 608 focused tests. An independent read-only
review finds no actionable defects; its fingerprint is recorded in the numeric
provenance. The new correctness cases reproduce and repair an empty
XML value that could prevent reader advancement when cached formula results
were disabled, and values truncated when split across text and CDATA. They
exercise numeric, string, boolean, shared-string, long-value, and typed-object
paths in addition to the measured APIs.

Use the [snapshot reproduction guide](excel-csv-broad-throughput-2026-10-04/reproduction/README.md)
with the baseline and candidate commits above, the candidate benchmark harness
in both snapshots, and this packet's [case file](excel-large-typed-read-2026-10-04/large-cases.json).
Select eight warmups, eight iterations, and one invocation for the buffer-guard
stage, or five warmups, five iterations, and one invocation for scan consolidation.
For the smaller-case follow-up, select the [three-case file](excel-large-typed-read-2026-10-04/pool-check-cases.json),
24 warmups, twelve iterations, and four invocations. The
[benchmark README](../../OfficeIMO.Excel.Benchmarks/README.md#large-typed-reads)
also describes direct full-scan comparisons and the OfficeIMO-only first-row lane.

For numeric decoding, use the candidate benchmark harness in both snapshots and
the [small-case file](excel-large-typed-read-2026-10-04/numeric-small-cases.json)
with 24/12/4 warmup/iteration/invocation counts. Use the
[large-case file](excel-large-typed-read-2026-10-04/numeric-large-measured-cases.json)
with 5/5/1 counts. The existing PowerForge rotated driver uses the
[seven-case file](excel-large-typed-read-2026-10-04/numeric-rotated-cases.json),
twelve warmups, 24 measured samples per engine, and four operations per sample
on each processor group. Resolve snapshot, case-file, and fixture paths to
absolute paths before invoking BenchmarkDotNet: its worker directory differs
from the caller's directory.

The packet retains [all native observations](excel-large-typed-read-2026-10-04/native.json),
[source and binary provenance](excel-large-typed-read-2026-10-04/provenance.json),
[fixture layout](excel-large-typed-read-2026-10-04/fixture-layout.json), and the
[allocation/thread-stack profile](excel-large-typed-read-2026-10-04/large-read-profile-summary.json).
