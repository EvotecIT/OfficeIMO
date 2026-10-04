# Large XLSX typed reads — 2026-10-04

Commit `cc694e167` avoids allocating a worksheet buffer when the package index
already declares that the worksheet exceeds the indexed reader's 64 MiB limit.
The reader uses its existing streaming fallback, preserving validation before
returning rows. The baseline is `9fb634005`, following the
[broader CSV/XLSX allocation changes](officeimo.excel-csv-broad-throughput-2026-10-04.md).

This removes about 128 MiB of managed allocation from opening the million-row
fixture. It does not resolve the remaining large-sheet allocation and opening
costs, and the contended timing run does not establish a speedup.

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

## Remaining allocation owners

A profile of three candidate million-row scans attributes roughly 190 MiB per
operation to used-range discovery and 198 MiB to worksheet validation. These
are weighted allocation samples, not exact accounting. Both scans inspect cell
references independently. Typed value decoding and row traversal account for
most of the remaining allocation. The profile's thread-stack counts must not be
interpreted as precise CPU time.

The [spreadsheet roadmap](../ROADMAP.md#spreadsheet-and-csv-delivery-order)
retains consolidation of those scans, large-sheet decoding, and portable timing
and memory qualification. The buffer change is one measured step toward those
targets.

## Validation and reproduction

Focused DataReader tests pass on Windows: 250 on .NET 10, 250 on .NET 8, and 245
on .NET Framework 4.7.2. The benchmark project builds for .NET 8 and .NET 10;
the product builds for `netstandard2.0`, with zero warnings or errors.
Before/after preflight compares complete generated row sets at every size and
the hash-pinned 65K corpus. The earlier Linux/WSL results cover the broader
implementation, not a new Linux run of this buffer decision.

Use the [snapshot reproduction guide](excel-csv-broad-throughput-2026-10-04/reproduction/README.md)
with the baseline and candidate commits above, the candidate benchmark harness
in both snapshots, and this packet's [case file](excel-large-typed-read-2026-10-04/large-cases.json).
Select eight warmups, eight iterations, and one invocation. The
[benchmark README](../../OfficeIMO.Excel.Benchmarks/README.md#large-typed-reads)
also describes direct full-scan comparisons and the OfficeIMO-only first-row lane.

The packet retains [all native observations](excel-large-typed-read-2026-10-04/native.json),
[source and binary provenance](excel-large-typed-read-2026-10-04/provenance.json),
[fixture layout](excel-large-typed-read-2026-10-04/fixture-layout.json), and the
[allocation/thread-stack profile](excel-large-typed-read-2026-10-04/large-read-profile-summary.json).
