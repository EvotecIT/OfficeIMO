# Wide high-cardinality CSV async read — 2026-09-28

Commit `3a92f3bed` added a public-reader workload with 32 fields and 5,000 or
25,000 rows. Every field value is distinct. `FirstRow` measures opening and
reading one row; `AllRows` awaits every row and obtains every field as a string.
The snapshot and incremental readers must return identical row, cell,
character, and signature totals. Setup checks every field of both readers
against the generated fixture before timing. File generation and validation
are outside timing, and file reads use the warmed operating-system cache.

BenchmarkDotNet 0.15.8 ran on Windows 11 build 26200, .NET 10.0.12, and an
AMD Ryzen 9 9950X3D2 with 32 logical processors. The two affinity masks
select separate 16-logical-processor domains. Both runs used High process
priority, six warmups, twelve measured iterations, and retained outliers.
The table reports means and allocated managed memory per operation.

| Affinity | Rows | Operation | Snapshot time | Incremental time | Snapshot allocation | Incremental allocation |
| --- | ---: | --- | ---: | ---: | ---: | ---: |
| `0xFFFF` | 5,000 | FirstRow | 11.294 ms | 0.227 ms | 22,331 KiB | 32.7 KiB |
| `0xFFFF` | 5,000 | AllRows | 11.358 ms | 11.191 ms | 22,332 KiB | 13,689 KiB |
| `0xFFFF` | 25,000 | FirstRow | 65.266 ms | 0.238 ms | 106,360 KiB | 32.7 KiB |
| `0xFFFF` | 25,000 | AllRows | 69.531 ms | 62.818 ms | 106,350 KiB | 68,331 KiB |
| `0xFFFF0000` | 5,000 | FirstRow | 9.515 ms | 0.223 ms | 22,331 KiB | 32.7 KiB |
| `0xFFFF0000` | 5,000 | AllRows | 10.002 ms | 12.824 ms | 22,333 KiB | 13,690 KiB |
| `0xFFFF0000` | 25,000 | FirstRow | 64.594 ms | 0.153 ms | 106,348 KiB | 32.7 KiB |
| `0xFFFF0000` | 25,000 | AllRows | 68.417 ms | 65.142 ms | 106,352 KiB | 68,332 KiB |

Incremental first-row access is consistently much faster and allocates only
about 33 KiB rather than materializing the file. Complete traversal allocates
about 36–39% less. Its elapsed-time result depends on size and processor
domain: the 5,000-row result is effectively tied on one domain and 28% slower
on the other, while the 25,000-row result is 4–9% faster. The small 25,000-row
time lead is within run variation and does not justify a parser change by
itself. This run does not establish a portable throughput ranking, peak
resident memory, cold-disk behavior, or a cross-platform regression budget.
The existing [sustained-read evidence](officeimo.csv-sustained-read-2026-09-27.md)
covers other reader and projection shapes; rotated comparisons and native
Linux/macOS qualification remain open in the [roadmap](../ROADMAP.md).

The runner commands were:

```powershell
./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload csvwideasyncread -RunMode full -Framework net10.0 -OutputRoot ./Ignore/Benchmarks/CsvWideAsync -AffinityMask 65535
./Build/Run-LibraryComparisonBenchmarks.ps1 -Workload csvwideasyncread -RunMode full -Framework net10.0 -OutputRoot ./Ignore/Benchmarks/CsvWideAsync -AffinityMask 4294901760
```

Raw logs, all measured samples, and normalized provenance results are retained
locally under `Ignore/Benchmarks/CsvWideAsync` for this active goal. The lane is
opt-in and does not update the website evidence catalog.
