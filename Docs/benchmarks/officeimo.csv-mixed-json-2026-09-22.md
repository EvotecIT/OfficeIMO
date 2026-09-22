# CSV mixed JSON file-write measurements — 2026-09-22

This run establishes a larger cross-platform file-write contract for short JSON
and mixed typed values. `CsvFileWriteBenchmarks.MixedJson25K` writes 25,000 rows
with a nullable JSON and Unicode text field, integers, decimals, UTC dates,
booleans, and a two-character `||` delimiter. It measures both `AsNeeded` and
`Always` quote modes.

Benchmark setup writes the complete file through OfficeIMO and CsvHelper, then
requires byte-for-byte identical UTF-8 output and validates every decoded field.
The `AsNeeded` file contains 3,127,872 bytes and the `Always` file contains
3,332,428 bytes on all three platforms.

## Measurement protocol

All runs use source commit `098c446a2acf96787d36eaaf8e2c74fde7b218bd`,
BenchmarkDotNet 0.15.8, .NET 10.0.12, 32 invocations per iteration, eight
warmups, 16 measured iterations, one launch, and retained outliers. Timing
includes file creation, serialization, UTF-8 encoding, and disposal through the
operating-system flush boundary. It excludes a durable-storage flush.

```powershell
dotnet run -c Release -f net10.0 --project OfficeIMO.CSV.Benchmarks -- --filter "*CsvFileWriteBenchmarks.*MixedJson25K*" --invocationCount 32 --unrollFactor 1 --warmupCount 8 --iterationCount 16 --launchCount 1 --outliers DontRemove
```

## Results

| Platform | Quote mode | OfficeIMO mean | OfficeIMO allocation | CsvHelper mean | CsvHelper allocation |
| --- | --- | ---: | ---: | ---: | ---: |
| Windows 11, AMD Ryzen 9 9950X3D2 | AsNeeded | 10.634 ms | 3.99 MB | 15.985 ms | 15.81 MB |
| Windows 11, AMD Ryzen 9 9950X3D2 | Always | 8.350 ms | 3.99 MB | 15.811 ms | 20.96 MB |
| Ubuntu 24.04, AMD Ryzen 9 9950X3D2 | AsNeeded | 17.08 ms | 3.99 MB | 30.84 ms | 15.81 MB |
| Ubuntu 24.04, AMD Ryzen 9 9950X3D2 | Always | 17.52 ms | 3.99 MB | 27.91 ms | 20.96 MB |
| macOS 27.0, Apple M4 | AsNeeded | 8.593 ms | 3.99 MB | 11.554 ms | 15.81 MB |
| macOS 27.0, Apple M4 | Always | 5.978 ms | 3.99 MB | 9.881 ms | 20.96 MB |

The hosts have different operating systems, processors, filesystems, and load,
so their absolute times are separate platform baselines. Several distributions
contained retained outliers, and the macOS `AsNeeded` OfficeIMO distribution was
bimodal. The allocation result repeated across all three platforms.

This evidence adds the larger corpus and Linux/macOS lanes needed for future
short JSON and mixed-field consistency work. It does not establish that the
remaining short-row timing variability is resolved.
