# Excel and CSV read/copy allocation — 2026-10-04

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
