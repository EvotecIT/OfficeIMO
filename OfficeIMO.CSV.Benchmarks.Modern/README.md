# Current CSV comparisons

I use this opt-in runner for CsvHelper 34.0.0 and FourLambda.Csv 1.1.0 alongside
OfficeIMO.CSV and Sep 0.17.1. CsvHelper 34 replaces the configuration and parser APIs used by
the existing comparison project. Keeping the runners separate preserves that
project's CsvHelper 33.1.0, wide-field, delimiter and writer comparisons. This
project stays outside the normal solution and does not ship with either library.
FourLambda's MPL-2.0 dependency is confined to this opt-in executable.

The 65K comparison opens the same hash-pinned sales CSV and returns owned strings
for every field. Each operation checks the fixed row/cell counts, decoded
character count and ordered checksum. Setup also validates every participant.
Four lanes use default string ownership. The fifth, `OfficeIMO_InternStrings`,
sets `InternStrings=true` and measures per-read string reuse while returning the
same complete observation. It is a separate option comparison; it is not the
default reader. None of these lanes measures borrowed spans. File data is warmed by setup, so this is not cold-disk I/O.
`--validate` checks the sales readers, generated owned-string readers, typed reads
and complete typed writes once, without collecting timings.

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.CSV.Benchmarks.Modern -- --filter "*MarkPflug65KModernCsvBenchmarks*" --inProcess --warmupCount 8 --iterationCount 12 --invocationCount 8 --unrollFactor 1 --outliers DontRemove
```


## Generated input

`CsvGeneratedModernReadBenchmarks` compares OfficeIMO, CsvHelper 34, Sep and
FourLambda on the same 50,000 mixed rows and ten columns used by the existing
[generated comparison](../OfficeIMO.CSV.Benchmarks/README.md#generated-excelreader-comparison).
Both projects use the shared record generator, header order and input serializer.
Each reader starts with identical UTF-8 bytes and returns an owned string for
every field. Setup prints the input SHA-256 and checks every decoded field,
header and row position against independently formatted expected values. Each
timed call verifies the row count and accumulated field lengths.

These methods use in-memory UTF-8 streams. OfficeIMO uses its default reader;
CsvHelper uses invariant culture, Sep uses comma separation with unescaping,
and FourLambda enables headers. The generated comparison in the original
project keeps its CsvHelper 33.1.0 API and additional read, parse and write lanes.
The shared OfficeIMO and Sep participants help compare the separate cohorts;
I keep their results separate rather than combining timings from different runs.

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.CSV.Benchmarks.Modern -- --filter "*CsvGeneratedModernReadBenchmarks*" --inProcess --warmupCount 8 --iterationCount 12 --invocationCount 8 --unrollFactor 1 --outliers DontRemove
```

## Typed reads and complete writes

`CsvGeneratedModernTypedBenchmarks` uses the same shared 50,000-row record generator
and ten-column schema. It compares OfficeIMO `RowsAs<T>` with CsvHelper 34's
source-generated record mapping. The write methods use `CsvDocument.WriteObjects`,
`CsvRowWriter.CreateStream` with `WriteObject`, and CsvHelper's source-generated
`WriteRecords`. Every write includes formatting, headers, final disposal, flushing
and copying the complete output into a byte array.

Setup validates every typed property and row position. Each writer's output is
read by both libraries, and its headers, byte count and SHA-256 are recorded.
Timed reads verify the full ten-property checksum and row count. The readers use
identical input bytes. Writers use their public default value formatting with
invariant culture and LF line endings, so equivalent decoded records can have
different output lengths. Stream buffer defaults are part of this complete-write
comparison. Date/time validation compares ticks. This catalog qualifies the JIT
object-mapping path, which carries the library's trimming annotation.

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.CSV.Benchmarks.Modern -- --filter "*CsvGeneratedModernTypedBenchmarks*" --inProcess --warmupCount 8 --iterationCount 12 --invocationCount 8 --unrollFactor 1 --outliers DontRemove
```

I retain the complete BenchmarkDotNet JSON, environment, package and source
identities, and all samples for a comparison. Different frameworks and machines
are separate cohorts. `OFFICEIMO_BENCHMARK_PROCESS_PRIORITY=Normal` selects the
same optional process-priority control as the other runners. CsvHelper is used under its Apache-2.0 license option;
Sep and BenchmarkDotNet are MIT. FourLambda is MPL-2.0 and uses its software
fallback on hosts without AVX2; I record that platform distinction with the results.
All comparison dependencies remain local to
this nonpackable executable.

## PackageAssets shapes

`PackageAssetsModernCsvBenchmarks` compares OfficeIMO, CsvHelper 34 and Sep on
the pinned 50,000-line NCsvPerf PackageAssets input in its `Plain`, `Quotes` and
`SpacesQuotes` shapes, described with the
[CsvHelper 33.1.0 lanes](../OfficeIMO.CSV.Benchmarks/README.md#packageassets-shapes).
Every reader returns an owned, unescaped string for each field of the same
in-memory UTF-8 bytes; `SpacesQuotes` trims whitespace outside quotes in all
three readers. Setup and every timed call check the counts, decoded characters
and ordered checksum against values computed from the source rows.

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.CSV.Benchmarks.Modern -- --filter "*PackageAssetsModernCsvBenchmarks*" --inProcess --warmupCount 6 --iterationCount 12 --invocationCount 8 --unrollFactor 1 --outliers DontRemove
```