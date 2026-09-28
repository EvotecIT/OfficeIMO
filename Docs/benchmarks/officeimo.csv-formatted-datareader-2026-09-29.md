# CSV formatted DataReader buffering — 2026-09-29

This measurement covers complete UTF-8 file export from a 25,000-row `IDataReader` with short JSON, Unicode text, integers, decimals, UTC dates, booleans, nulls, and a `||` delimiter. OfficeIMO now batches rows on this formatted, multi-character-delimiter path. The ordinary short-text path stays unbatched. The writer also commits completed buffered rows if a reader, formatter, or cancellation fails, without writing an incomplete current row; the same failure rule applies to the existing typed DataReader batch path.

`CsvFileWriteBenchmarks.MixedJson25K` creates and closes each file inside the timed operation. Setup requires byte-identical OfficeIMO and CsvHelper output and validates every decoded field. Each successful `AsNeeded` file has 3,127,872 bytes; `Always` has 3,332,428 bytes. The comparison includes operating-system flush on disposal, not a durable-storage flush.

The source baseline was `56435017597bb37c0ad8f495a9738f4a749a3e83`. All runs used Windows 11, an AMD Ryzen 9 9950X3D2, .NET 10.0.12, BenchmarkDotNet 0.15.8, Normal process priority, 16 invocations per iteration, eight warmups, 12 measured iterations, one launch, and retained outliers. Each run measured both 16-logical-processor cache domains separately. Baseline and candidate runs were rotated; the ranges below include every run of the final source shape and all three baseline controls.

| Quote mode | Cache domain | Baseline mean range | Candidate mean range | Allocation, baseline → candidate |
| --- | --- | ---: | ---: | ---: |
| AsNeeded | `0xFFFF` | 8.139–8.300 ms | 6.192–8.333 ms | 387.37 → 436.72 KB |
| AsNeeded | `0xFFFF0000` | 6.081–8.342 ms | 5.768–6.317 ms | 387.37 → 436.72 KB |
| Always | `0xFFFF` | 6.974–7.898 ms | 6.183–6.304 ms | 387.31 → 436.67 KB |
| Always | `0xFFFF0000` | 5.839–6.182 ms | 5.857–6.639 ms | 387.31 → 436.67 KB |

The candidate usually took less time, but its runs overlap the baseline and one `Always`/`0xFFFF0000` run was slower. The comparison writer also moved substantially between runs, so this evidence does not establish a repeatable elapsed-time reduction. The candidate allocates about 49 KB more per export. A broader first attempt batched the 1,000-row short Unicode case as well; it raised allocation without a consistent time gain, so the implementation limits batching to formatted multi-character-delimiter readers. The CSV write-consistency and portable budget work remains open.

Reproduce the file lane from the same source and runtime with:

```powershell
$env:OFFICEIMO_BENCHMARK_OUTPUT = Join-Path (Get-Location) 'Ignore/Benchmarks/csv-formatted-files'
dotnet run -c Release -f net10.0 --project ./OfficeIMO.CSV.Benchmarks -- --filter '*CsvFileWriteBenchmarks.*MixedJson25K*' --priority Normal --affinityMasks 0xFFFF,0xFFFF0000 --invocationCount 16 --unrollFactor 1 --warmupCount 8 --iterationCount 12 --launchCount 1 --outliers DontRemove --artifacts ./Ignore/Benchmarks/csv-formatted-results
```

The cache-domain masks are specific to this workstation. Recheck processor topology before using them elsewhere. The failure-path contract was checked with cancellation and formatter exceptions on .NET 8 and .NET Framework 4.7.2; the full .NET 8 CSV suite and .NET Standard 2.0 build also passed. These tests do not establish Linux or macOS performance budgets.
