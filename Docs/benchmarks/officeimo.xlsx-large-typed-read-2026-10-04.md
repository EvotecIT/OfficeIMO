# Large XLSX typed reads — 2026-10-04

Large raw packets are retained in the [measurement archive](excel-large-typed-read-2026-10-04/raw-packets.zip), with original filenames and exact bytes recorded in its [SHA-256 manifest](excel-large-typed-read-2026-10-04/raw-packets-manifest.json). Links to those packets download the archive. [Extraction instructions](excel-large-typed-read-2026-10-04/raw-packets.README.md) explain how to inspect the individual captures; smaller summaries and reproduction sources remain beside it.

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
Both the [initial observations](excel-large-typed-read-2026-10-04/raw-packets.zip)
and [longer-warmup observations](excel-large-typed-read-2026-10-04/raw-packets.zip)
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

All [30 native cases](excel-large-typed-read-2026-10-04/raw-packets.zip)
complete. Timing remains mixed. A follow-up rotates all six numeric cases and
the 65K control across both verified processor groups, retaining
[672 measured samples](excel-large-typed-read-2026-10-04/raw-packets.zip).
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

All [38 native cases](excel-large-typed-read-2026-10-04/raw-packets.zip)
complete. Ordinary and prefixed shared-string controls have effectively unchanged
allocation. The 65K corpus again varies with pool state: the baseline reports
2.57 MiB in this run versus about 90 KiB in earlier longer-warmup observations.
That difference is not credited to coordinate buffering.

The allocation improvement does **not** qualify the initial helper as a
throughput improvement. The
[15-scenario rotated comparison](excel-large-typed-read-2026-10-04/raw-packets.zip)
retains 1,440 samples, including slower materialized and numeric reads. A separate
[large-sheet rotation](excel-large-typed-read-2026-10-04/raw-packets.zip)
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

All [46 native cases](excel-large-typed-read-2026-10-04/raw-packets.zip)
and 2,016 output-validated rotated samples complete. The smaller workloads use
24 native warmups, twelve iterations, and four operations per iteration; the
large workloads use five warmups, five iterations, and one operation. PowerForge
rotates both versions on each processor group with the settings recorded in the
[provenance](excel-large-typed-read-2026-10-04/typed-values-provenance.json).
Every sample and outlier remains available in the
[small](excel-large-typed-read-2026-10-04/raw-packets.zip) and
[large](excel-large-typed-read-2026-10-04/raw-packets.zip) packets.

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
[measurements](excel-large-typed-read-2026-10-04/raw-packets.zip)
and [patch](excel-large-typed-read-2026-10-04/coordinate-rejected-inline.patch)
remain for reproduction. Separate
[validation traversal](excel-large-typed-read-2026-10-04/raw-packets.zip)
and [schema-name](excel-large-typed-read-2026-10-04/raw-packets.zip)
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
[native packet](excel-large-typed-read-2026-10-04/raw-packets.zip)
contains twelve cases, and the
[rotated packet](excel-large-typed-read-2026-10-04/raw-packets.zip)
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

## Large-file peer refresh after coordinate correction

The [18-case refresh](excel-large-typed-read-2026-10-04/raw-packets.zip)
uses source `15b5f2848`, Sylvan.Data.Excel 0.5.8, and ExcelReader.NET 5.1.1.
Every engine consumes the same four typed fields, with all rows and headers
validated before measurement. Five warmups and five measured invocations run
at Normal priority on each processor group; outliers are retained. These are
sequential BenchmarkDotNet cases, so this diagnostic refresh does not replace
rotated ordering or portable qualification.

| Rows | OfficeIMO mean ms, groups A / B | Sylvan mean ms, A / B | ExcelReader.NET mean ms, A / B | Allocated MiB, OfficeIMO / Sylvan / ExcelReader.NET |
| --- | ---: | ---: | ---: | ---: |
| 25,000 | 127.9 / 114.7 | 99.8 / 79.1 | 27.9 / 23.7 | 1.472 / 0.336 / 0.017 |
| 250,000 | 709.0 / 828.9 | 409.4 / 243.0 | 93.7 / 66.5 | 23.541 / 0.456 / 0.017 |
| 1,000,000 | 2,640.9 / 2,188.0 | 952.0 / 862.4 | 337.4 / 469.4 | 93.446 / 0.866 / 0.017 |

Groups A and B use masks `0xFFFF` and `0xFFFF0000`, respectively. The packet
retains every observation and the source, binary, harness, runtime, and host
identity. Substantial throughput and allocation gaps remain at this checkpoint.
The earlier allocation reductions do not establish competitive leadership.
OfficeIMO's eager validation is part of its measured full-read cost; peer
first-row behavior is not treated as an equivalent contract.

## XML type and style attributes

Commit `33d26b66d` reads ordinary cell type and style attributes into a reusable
character buffer instead of allocating a string for each attribute. Known cell
types reuse constants, and a parsed style index supplies both date classification
and the 1904-calendar adjustment. Long, escaped, empty, and unknown attributes
retain the XML reader's value semantics. The baseline is `15b5f2848`; the benchmark
harness is identical in both binary snapshots.

The warmed native matrix contains 54 before/after cases. PowerForge contributes
2,400 successful rotated samples across both processor groups. This stage extends
the allocation result to multiple sizes and materialized APIs:

| Workload | Before allocation | After allocation |
| --- | ---: | ---: |
| 1,000,000 rows, all typed fields | 93.447 MiB | 2.008 MiB |
| 1,000,000 rows, first-row operation | 46.866 MiB | 1.147 MiB |
| 250,000 rows, all typed fields | 23.541 MiB | 0.680 MiB |
| 250,000 rows, first-row operation | 11.961 MiB | 0.483 MiB |
| 25,000 rows, all typed fields | 1.471 MiB | 0.327 MiB |
| 25,000 rows, first-row operation | 1.471 MiB | 0.327 MiB |
| 2,500-row materialized DataTable | 1.833 MiB | 1.547 MiB |
| 2,500-row materialized objects | 2.333 MiB | 2.048 MiB |

Numeric XML without type/style attributes, omitted-coordinate cases, and ordinary
or prefixed shared-string cases have essentially unchanged allocation. The 65K
control measures 2.572 MiB in both snapshots in this run; that is higher than the
earlier warmed observations and does not establish a stable absolute budget.
The object-valued numeric DataReader cases vary by roughly 16 KiB in opposite
directions. Keep those observations alongside the improvements.

Throughput remains unresolved. The table shows candidate/baseline median ratios
from the rotated runs; values above 1 are slower. Means and every raw sample are
retained, including large timing excursions while other workstation work ran.

| Workload | First processor group | Second processor group |
| --- | ---: | ---: |
| 1,000,000 rows, all typed fields | 1.026 | 1.043 |
| 1,000,000 rows, first-row operation | 1.019 | 1.047 |
| 250,000 rows, all typed fields | 1.049 | 1.078 |
| 250,000 rows, first-row operation | 1.053 | 1.053 |
| 25,000 rows, all typed fields | 1.013 | 1.123 |
| 25,000 rows, first-row operation | 1.088 | 1.132 |
| Numeric XML DataTable, double values | 1.265 | 1.282 |
| Numeric XML DataTable, decimal values | 1.067 | 1.089 |
| Materialized objects | 1.073 | 1.141 |

These slower cases remain performance work. This stage qualifies a substantial
allocation reduction, not a general speedup or competitive lead. Allocation also
does not measure retained or peak process memory.

The new metadata tests expose a pre-existing date error: a 1904 workbook using
the valid padded/signed style index `" +1 "` returned `2024-02-29` instead of
`2028-03-01`. The saved baseline reproduces the failure. Using the same parsed
style index for both date decisions fixes it across typed getters, object values,
ranges, DataTable, and typed objects. Tests also cover numeric entities, long
style indices, unknown types spanning a surrogate boundary, invalid styles, and
cell-reference diagnostics.

The full .NET 10 suite passes 5,376 tests with five skips. Focused suites pass 760
tests on .NET 8, 755 on .NET Framework 4.7.2, and 760 on Ubuntu/WSL .NET 10. The
`netstandard2.0` product builds without warnings or errors. An independent
read-only review reports no actionable findings. These checks qualify correctness
on those runtimes; portable timing and memory budgets remain open.

Reproduce the large cases with 5/5/1 warmup/iteration/invocation counts and the
[small-case matrix](excel-large-typed-read-2026-10-04/metadata-small-cases.json)
with 24/12/4 counts. Rotated large runs use 3/12/1; small runs use 12/24/4. All
runs retain outliers and use Normal priority with the two previously recorded
processor masks. The packet contains [native results](excel-large-typed-read-2026-10-04/raw-packets.zip),
[rotated observations](excel-large-typed-read-2026-10-04/raw-packets.zip),
and [source, binary, review, and test provenance](excel-large-typed-read-2026-10-04/metadata-attributes-provenance.json).
The [refreshed profile](excel-large-typed-read-2026-10-04/metadata-attributes-profile.json)
records 99 weighted allocation samples across five complete million-row scans,
excluding fixture generation. XML parsing and package reads remain investigation
targets; sampled thread stacks do not establish precise CPU-time attribution.

An earlier 128 KiB worksheet-stream buffering experiment was removed. Its eight
native cases show inconsistent timing changes and additional allocation. The
[rejected results and original patch text](excel-large-typed-read-2026-10-04/raw-packets.zip)
remain reproducible evidence, not part of the product implementation.

## Reusing inferred coordinates from complete scans

Commit `e7680caf1` reuses inferred row coordinates from the complete validation
scan. Used-range discovery also caches the empty index for sequential sheets.
Sparse discovery keeps coordinate-map construction lazy, so callers requesting
only a range do not retain a map for rows they never read. Publication follows
the complete XML scan and its final cancellation check.

The output-validated matrix covers 28 workloads at 2,500 and 25,000 rows:
explicit coordinates, omitted row coordinates, omitted row and cell coordinates,
typed readers, explicit and discovered ranges, discovery alone on dense and
sparse sheets, and double/decimal DataTable controls. All 56 native cases and
2,688 rotated samples complete successfully. The native comparison uses
24 warmups, 12 measurements and eight operations; rotated runs use 12 warmups,
24 measurements and four operations on each processor group.

| 25,000-row workload | Median ratio, group A / B | Managed allocation before / after |
| --- | ---: | ---: |
| Omitted row coordinates, typed reader | 0.75 / 0.79 | 313,919 / 292,731 B |
| Omitted row coordinates, discovered range | 0.77 / 0.77 | 2,282,652 / 2,261,396 B |
| Omitted row and cell coordinates, typed reader | 0.78 / 0.78 | 297,643 / 282,057 B |
| Omitted row and cell coordinates, discovered range | 0.80 / 0.78 | 2,266,498 / 2,250,682 B |
| Sparse omitted coordinates, discovery only | 1.00 / 1.00 | 225,328 / 225,922 B |

A ratio below one favors the candidate. The four omitted-coordinate projection
cases also improve at 2,500 rows, with ratios of 0.76–0.92 across the two groups.
Explicit typed-reader and discovered-range controls remain within 3%. Other
controls are less stable: dense omitted-coordinate discovery measures 2–13%
slower at 25,000 rows, and decimal DataTable materialization measures 7–11%
slower despite unchanged allocation. These signals remain open; the results
do not establish a universal speedup on this busy workstation.

The first eager-cache design increased 25,000-row sparse discovery allocation
from 225,832 to 2,185,824 bytes. It was rejected. Its
[native and valid rotated observations](excel-large-typed-read-2026-10-04/raw-packets.zip)
remain separate from the final candidate. A preliminary rotation invalidated
by a source-provenance change contributes no accepted timing samples.

The final candidate passes 5,376 Windows .NET 10 tests with five existing skips,
760 focused .NET 8 tests, 755 .NET Framework 4.7.2 tests, and 760 Linux/WSL .NET 10
tests. The product builds for `netstandard2.0`; benchmarks build for .NET 8 and
.NET 10. Independent review and one targeted confirmation report no actionable
findings. The packet retains [native observations](excel-large-typed-read-2026-10-04/raw-packets.zip),
[rotated samples](excel-large-typed-read-2026-10-04/raw-packets.zip),
[case definitions](excel-large-typed-read-2026-10-04/coordinate-cache-cases.json),
and [source, binary, review and validation provenance](excel-large-typed-read-2026-10-04/coordinate-cache-provenance.json).
Warmed allocation does not measure retained or peak memory, and correctness
on WSL does not qualify native-Linux timing.

## Rejected larger XML-reader buffers

An experiment enabling the framework XML reader's asynchronous configuration
while continuing synchronous reads uses its larger internal buffers. It is
removed after an eight-workload, 16-case native screen. The million-row full scan
allocates 2,105,480 → 959,464 bytes, but its median is slower; first-row medians
are slower at both 250,000 and 1,000,000 rows. The 25,000-row full scan allocates
another 227,538 bytes and also measures slower. Numeric DataTable medians improve
in this screen but allocate another 165 KiB. The 65K control remains 91,771 bytes.

The candidate passes 760 focused .NET 10 reader tests and complete output
preflight for all eight workloads. These results do not qualify a general gain;
the [screen packet](excel-large-typed-read-2026-10-04/raw-packets.zip)
retains every observation, source patch, binary hash and job setting. The
integrated baseline at `a97198615` passes 5,376 Excel tests with five skips and
638 CSV tests on Windows .NET 10. No rotated or other-runtime qualification is
claimed for the rejected buffer change.

## Short XML attribute completion

The XML attribute readers avoid a second `ReadValueChunk` call when the framework
reader has already returned a complete short value. A chunk of 31 or 32 characters
still probes for continuation because a surrogate pair can cross the boundary.
Long-value handling and restoration of the reader's element position are unchanged.

The qualification covers 42 workloads: 250,000- and 1,000,000-row first/full reads,
25,000-row reads, the 65K corpus, typed and materialized numeric APIs, shared
strings, explicit and inferred coordinates, and dense/sparse range discovery.
All 84 native cases and 3,840 rotated samples pass their output checks. Snapshots
use identical benchmark and dependency DLLs; only the Excel product DLL differs.

Million-row full-read median ratios are 0.954/0.944 on the two processor groups;
first-row ratios are 0.962/0.925. The 250,000-row full-read ratios are 0.917/0.974,
and first-row ratios are 0.922/0.956. Managed allocation is effectively unchanged:
the million-row full scan measures 2,105,480 → 2,105,560 bytes. Native sequential
jobs show slower large-read medians, and the group-A 250,000-row first-read mean
is 4.9% slower despite a lower median. All observations remain in the packet;
these results qualify a bounded large-read improvement, not a portable ranking.

An initial 2,500-row explicit-range slowdown prompted a longer four-workload
follow-up with identical-baseline controls. It retains another 1,536 samples,
using 24 warmups, 48 measurements and 16 operations per sample. Explicit-range
candidate medians are about 9% lower on both groups in that follow-up. The
identical-baseline ratios range from 0.94 to 1.03, demonstrating the host's
timing variability. Small omitted-coordinate used-range results remain mixed:
the follow-up measures 8% slower on one group and 2% faster on the other. The
small-workload measurements do not establish a general speedup or close the
earlier numeric materialization targets.

Correctness passes 5,376 Windows .NET 10 tests with five existing skips, 760
focused .NET 8 tests, 755 .NET Framework 4.7.2 tests, 760 Linux/WSL .NET 10 tests,
and 760 macOS ARM64 .NET 10 tests. Product `netstandard2.0` and benchmark
.NET 8/.NET 10 builds succeed. Independent read-only review reports no actionable
findings. The [native cases](excel-large-typed-read-2026-10-04/raw-packets.zip),
[full rotations](excel-large-typed-read-2026-10-04/raw-packets.zip),
[focused follow-up and controls](excel-large-typed-read-2026-10-04/raw-packets.zip),
and [source, binary, runtime and reproduction details](excel-large-typed-read-2026-10-04/raw-packets.zip)
retain the complete evidence. These runs measure warmed managed allocation;
they do not measure retained or peak memory.

## Earlier allocation profile and remaining work

A profile after the buffer guard, before scan consolidation, attributes roughly
190 MiB per operation to used-range discovery and 198 MiB to worksheet validation
across three million-row scans. These
are weighted allocation samples, not exact accounting. Both scans inspect cell
references independently. Typed value decoding and row traversal account for
most of the remaining allocation. The profile's thread-stack counts must not be
interpreted as precise CPU time.

The [spreadsheet roadmap](../ROADMAP.md#spreadsheet-and-csv-delivery-order)
retains large-sheet decoding, general read/export throughput, and portable timing
and memory qualification. The omitted-row coordinate disagreement found during
scan consolidation is fixed by the coordinate stage above. The coordinate-cache
stage recovers part of its additional scan cost; general XML-reader throughput
and the unfavorable controls remain open.

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

The packet retains [all native observations](excel-large-typed-read-2026-10-04/raw-packets.zip),
[source and binary provenance](excel-large-typed-read-2026-10-04/provenance.json),
[fixture layout](excel-large-typed-read-2026-10-04/fixture-layout.json), and the
[allocation/thread-stack profile](excel-large-typed-read-2026-10-04/large-read-profile-summary.json).
