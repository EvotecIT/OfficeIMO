# Excel and CSV throughput and allocation — 2026-10-04

The changes in `fe9daeb03` reduce warm managed allocation for short JSON CSV
exports by 44.5% and for small XLSX shared-string tables with a Unicode or XML
entity tail by about 11.6%. The comparison baseline is `dfb5753794`.

This run does not establish a general speedup or an overall library ranking.
The workstation reached 100% CPU utilization during measurement, with other
builds and applications active. Even identical-version controls produced large
timing differences. All measured samples are retained in the evidence packet.

## Implementation and observed allocation

| Workload | Before | After | Interpretation |
| --- | ---: | ---: | --- |
| CSV: 1,000 short JSON rows, quote as needed | 346.18 KiB | 192.16 KiB | 44.5% less allocation |
| XLSX: 256 shared strings, final Unicode value | 103.04–103.12 KiB | 91.03–91.10 KiB | About 11.6% less allocation |
| XLSX: 256 shared strings, final entity-escaped value | 103.04–103.12 KiB | 91.03–91.10 KiB | About 11.6% less allocation |
| XLSX: 256 ordinary ASCII shared strings | 76.35–76.38 KiB | 76.31–76.38 KiB | Effectively unchanged |

These are BenchmarkDotNet managed bytes per completed operation after warmup.
They exclude cold pool acquisition, retained pool capacity, process working set,
and peak memory. CSV output remains identical to the validated reference text;
the Excel read changes do not write or alter workbooks.

`CsvRowWriter.WriteDataReader` accumulates completed rows from text-only readers
in a rented character buffer instead of repeatedly growing and clearing a large
`StringBuilder`. The pooled default path batches about 32 Ki characters; the existing formatted-text
path retains its smaller threshold. The buffer is returned with clearing.
Separating the current row from completed rows also fixes the legacy fallback
case where a formatter failure emitted part of the failed row. Earlier complete
rows remain available. Modern readers with numeric or other non-string field
types retain the established direct `StringBuilder` batch: the initial pooled
version saved only about 2% allocation on the 25K quoted workload and showed a
26% slowdown on one CPU domain. The final change restores that path rather than
accepting the tradeoff. The reader implementation lives in its own partial file.

An experiment routing short densely quoted text through the chunk escape helper
was also removed. Its setup cost affected small quoted fields in the typed
workload. The final escape implementation is identical to the
baseline source, including its existing long-field chunk path and quote-run
optimization.

The Excel small-part ASCII reader checks for non-ASCII bytes and entity markers
before allocating decoded prefix strings. The existing full XML reader handles
those inputs directly. The check is confined to the existing pooled-part path,
whose part limit is 8,192 bytes. XML support, resource limits, cancellation,
fallback behavior, public APIs, and production dependencies are unchanged.

## Equivalent competitor comparisons

The following ranges are the two CPU-domain means, not confidence intervals.
They describe this run under contention. Compare libraries within a row; do not
compare these absolute times with the separately hosted before/after protocol.

| Completed operation | OfficeIMO time / allocation | Peer time / allocation |
| --- | --- | --- |
| 1,000 short JSON CSV rows, quote as needed | 92.2–93.6 µs / 192.16 KiB | CsvHelper: 275.8–282.9 µs / 618.09 KiB |
| 25,000 mixed CSV rows through `IDataReader`, sequential | 4.98–17.95 ms / 5.51 MiB | Sylvan: 3.96–21.04 ms / 5.45 MiB |
| 65K CSV rows, materialize and consume every field as a string | 20.03–22.72 ms / 34.21 MiB | Sep: 21.69–25.15 ms / 34.22 MiB; Sylvan: 21.69–24.20 ms / 35.75 MiB |
| 256 XLSX shared strings, Unicode tail | 341.8–413.4 µs / 91.17–91.64 KiB | Sylvan: 528.0–690.9 µs / 315.65–315.92 KiB |
| 65K XLSX rows, typed scan of fourteen columns | 89.14–113.34 ms / 89.33–2,633.44 KiB | ExcelReader.NET: 90.51–105.99 ms / 31.41 KiB; Sylvan: 201.40–243.46 ms / 648.98 KiB |
| 1,000 long plain-text XLSX rows, compact package | 2.94–3.21 ms / 246.14 KiB | SpreadCheetah: 3.34–6.96 ms / 224.76 KiB |

OfficeIMO's short JSON allocation is about 69% below CsvHelper's in this lane.
The string-materializing CSV scan is close to Sep's allocation floor; the
[separate span-reader evidence](officeimo.csv-wide-read-dispatch-2026-09-28.md)
addresses a different result contract. ExcelReader.NET retains a material
allocation advantage on the typed XLSX scan. SpreadCheetah allocates less on
the compact long-text export. The mixed CSV writer's timing ranking reverses
between CPU domains; those times cannot establish a winner. The observed long-text
XLSX elapsed-time lead
does not close the [long-plain throughput target](officeimo.xlsx-long-text-2026-09-28.md).

The first-domain OfficeIMO XLSX scan allocated 2,696,640 bytes per operation,
versus 91,470 on the second domain. That higher observation is included above
and in the raw evidence; it is not replaced by the lower warmed result.
A separate 32-scan warmed allocation diagnostic reports 91,567 bytes before and
91,529 after, so it does not reproduce that spike or show a material change in
this large-workbook path. The spike's cause remains unqualified.

The new 256-row XLSX fixture is synthetic. Both libraries read the exact same
package, and setup checks every field and row count. It isolates late shared-
string fallback rather than representing all Excel-produced workbooks. The 65K
CSV and XLSX fixtures come from the pinned MarkPflug benchmark corpus, with
SHA-256 checks and complete observation validation. CSV writer setup validates
decoded fields; the short-JSON writer comparison additionally requires identical
text. XLSX export setup reopens and validates every value with ExcelDataReader.

## Timing controls and remaining qualification

The initial rotated runs retain 24 samples per engine with 12 warmups. Follow-up
controls retain 32 samples with 16 warmups. PowerForge alternates the order,
and no outliers are discarded. Identical baseline assemblies are loaded for
both engines in the A/A controls. Those controls show 3.21–3.45x median
differences for short JSON with always-quote, so this hosted timing path cannot
certify that workload. The candidate's short JSON/as-needed follow-up medians
are 0.84–0.87x baseline, but that encouraging result still needs quiet-host proof.

For the initial candidate, the 25K quoted-row follow-up is 1.28x baseline median on the first domain and
1.03x on the second. This is an unresolved throughput signal, not a qualified
speed improvement. A separate-process check also showed a 26% slowdown on one
domain. This evidence motivated restoring direct typed batching in the final
candidate. The short-field escape experiment was subsequently removed as well.
All controls and superseded-candidate results remain in the packet,
labelled by stage. The ordinary ASCII XLSX follow-up stays within about 2% of baseline.

The final separate-process 25K quoted scan measures 6.05 → 5.53 ms and
7.33 → 5.56 ms across the two domains, with unchanged 7.09 MiB allocation.
The earlier typed-export slowdown does not recur in that run. Final short JSON
means are 87.20 → 93.58 µs and 69.93 → 74.85 µs, about 7–8% higher, with
overlapping confidence intervals. Its 44.5% allocation reduction is the qualified
gain; throughput remains an open qualification rather than a promised speedup.
Long JSON means move in opposite directions across domains, while allocation
drops from 10.24 to 10.11 MiB. These measurements do not justify a universal
performance claim.

Namespace-prefixed worksheet elements triggered the existing SDK fallback during
fixture construction. That exploratory fixture was replaced with a default-
namespace worksheet before collecting the shared-string measurements above.
Its exploratory timings are excluded because they measured a different reader
path. Supporting the prefixed form in the native reader requires its own
namespace and format qualification.

Remaining work lives in the [product roadmap](../ROADMAP.md): qualify throughput
on a quiet host, investigate allocation variance and retained/peak memory on
large workbooks, broaden native shared-string and namespace handling, and add
native Linux/macOS evidence. This run does not establish portable budgets.

## Validation and reproduction

- CSV correctness: 628 tests on .NET 10, 628 on .NET 8, and 440 on .NET Framework 4.7.2.
- Excel DataReader/shared-string correctness: 286 tests on .NET 10, 286 on .NET 8, and 281 on .NET Framework 4.7.2.
- Both product projects build for `netstandard2.0` with zero warnings or errors.
- One independent read-only review and one targeted confirmation of restored typed batching found no actionable P0–P3 defects. They covered row failure/cancellation, destination exceptions, pooled ownership, legacy fallback, escaping, XML limits, and benchmark equivalence. The final escape file is restored to baseline. The review notes that `GetFieldType` throwing on the modern default path is not directly covered by the changed regression helper; legacy and formatted-reader tests exercise the shared value-buffering path.

The evidence packet contains [native samples](excel-csv-throughput-2026-10-04/native.json),
[rotated samples and controls](excel-csv-throughput-2026-10-04/rotated.json),
[source and assembly provenance](excel-csv-throughput-2026-10-04/provenance.json),
and [validation results](excel-csv-throughput-2026-10-04/validation.json).
The [reproduction instructions](excel-csv-throughput-2026-10-04/reproduction/README.md)
include the isolated snapshot loader and existing PowerForge driver.
It retains 122 successful native benchmark cases and twelve rotated/control
runs, including the intermediate implementations that were narrowed or removed.

BenchmarkDotNet 0.15.8 used .NET 10.0.12 and SDK 10.0.112 on Windows 11
10.0.26300.9457, Ryzen 9 9950X3D2, Normal process priority, and the High
performance power plan. Windows topology inspection reported two 96 MiB L3
domains, selected by masks `0xFFFF` and `0xFFFF0000`. These masks are specific
to the measured workstation. Native comparisons retain twelve measured
iterations after six warmups, with one launch and no outlier removal; the
additional large before/after check retains sixteen after eight warmups.

The tested peer versions are CsvHelper 33.1.0, Sep 0.17.0, Sylvan.Data.Csv 1.4.4,
Sylvan.Data.Excel 0.5.8, ExcelReader.NET 3.0.1, and SpreadCheetah 1.28.0. They are
the repository's existing benchmark pins, not a claim that each is the latest
available version. Existing test-only dependency and license boundaries remain
in place; no production dependency is added.
