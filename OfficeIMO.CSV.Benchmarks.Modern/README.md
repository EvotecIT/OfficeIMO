# Current CSV reader comparisons

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
`--validate` runs the five sales-reader configurations and the four generated-data
readers once, without collecting timings.

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

I retain the complete BenchmarkDotNet JSON, environment, package and source
identities, and all samples for a comparison. Different frameworks and machines
are separate cohorts. `OFFICEIMO_BENCHMARK_PROCESS_PRIORITY=Normal` selects the
same optional process-priority control as the other runners. CsvHelper is used under its Apache-2.0 license option;
Sep and BenchmarkDotNet are MIT. FourLambda is MPL-2.0 and uses its software
fallback on hosts without AVX2; I record that platform distinction with the results.
All comparison dependencies remain local to
this nonpackable executable.
