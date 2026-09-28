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
An Ubuntu 24.04.3 run under WSL on the same host and source commit `9fc03dd74`
measured the same eight cases with .NET 10.0.12. It used `taskset` to pin the
parent process to logical processors 0–15 or 16–31; benchmark setup checked
the inherited worker mask and every generated field. Workers ran at Normal
priority. Six warmups, twelve measured iterations, and retained outliers
matched the Windows measurement policy. The table reports means and managed
allocation per operation; the two domains were measured in A-then-B order.

| WSL affinity | Rows | Operation | Snapshot time | Incremental time | Snapshot allocation | Incremental allocation |
| --- | ---: | --- | ---: | ---: | ---: | ---: |
| `0xFFFF` | 5,000 | FirstRow | 12.572 ms | 0.082 ms | 22,305 KiB | 32.4 KiB |
| `0xFFFF` | 5,000 | AllRows | 9.927 ms | 10.272 ms | 22,305 KiB | 13,641 KiB |
| `0xFFFF` | 25,000 | FirstRow | 104.554 ms | 0.128 ms | 106,218 KiB | 32.4 KiB |
| `0xFFFF` | 25,000 | AllRows | 80.090 ms | 38.821 ms | 106,219 KiB | 68,084 KiB |
| `0xFFFF0000` | 5,000 | FirstRow | 12.988 ms | 0.056 ms | 22,306 KiB | 32.4 KiB |
| `0xFFFF0000` | 5,000 | AllRows | 11.067 ms | 11.498 ms | 22,305 KiB | 13,640 KiB |
| `0xFFFF0000` | 25,000 | FirstRow | 84.900 ms | 0.094 ms | 106,219 KiB | 32.4 KiB |
| `0xFFFF0000` | 25,000 | AllRows | 83.036 ms | 37.358 ms | 106,218 KiB | 68,086 KiB |

First-row allocation and the 36–39% complete-traversal allocation reduction
replicate the Windows pattern. Complete-traversal time is mixed at 5,000 rows
and lower for the incremental reader at 25,000 rows. Timing variation is
material: the 5,000-row incremental AllRows 99.9% interval half-width is
6.2 ms on A and 6.1 ms on B. WSL uses the Windows-backed worktree here, so
these are cross-runtime observations rather than native-Linux throughput or
portable latency budgets. The existing [sustained-read evidence](officeimo.csv-sustained-read-2026-09-27.md)
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

The WSL sample logs and BenchmarkDotNet reports are in the
`wsl-ubuntu2404-20260928-domain-a` and `wsl-ubuntu2404-20260928-domain-b`
subdirectories of that same ignored folder. From a WSL shell in the repository
root, this reproduces the lane; use CPU range `16-31`, mask `0xFFFF0000`, and
the domain-b output suffix for the second domain:

```sh
dotnet build -c Release -f net10.0 \
  OfficeIMO.CSV.Benchmarks/OfficeIMO.CSV.Benchmarks.csproj
taskset -c 0-15 env OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY=0xFFFF \
  dotnet run -c Release -f net10.0 --no-build \
  --project OfficeIMO.CSV.Benchmarks/OfficeIMO.CSV.Benchmarks.csproj -- \
  --filter '*CsvWideAsyncReadBenchmarks*' \
  --artifacts Ignore/Benchmarks/CsvWideAsync/wsl-ubuntu2404-20260928-domain-a \
  --warmupCount 6 --iterationCount 12 --outliers DontRemove
```
