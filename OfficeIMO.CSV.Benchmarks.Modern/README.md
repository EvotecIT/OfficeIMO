# Current CSV reader comparisons

I use this opt-in runner for CsvHelper 34.0.0 and FourLambda.Csv 1.1.0 alongside
OfficeIMO.CSV and Sep 0.17.1. CsvHelper 34 replaces the configuration and parser APIs used by
the existing comparison project. Keeping the runners separate preserves that
project's CsvHelper 33.1.0, wide-field, delimiter and writer comparisons. This
project stays outside the normal solution and does not ship with either library.
FourLambda's MPL-2.0 dependency is confined to this opt-in executable.

Every reader opens the same hash-pinned 65K sales CSV and returns owned strings
for every field. Each operation checks the fixed row/cell counts, decoded
character count and ordered checksum. Setup also validates every participant.
Four lanes use default string ownership. The fifth, `OfficeIMO_InternStrings`,
sets `InternStrings=true` and measures per-read string reuse while returning the
same complete observation. It is a separate option comparison; it is not the
default reader. None of these lanes measures borrowed spans. File data is warmed by setup, so this is not cold-disk I/O.
`--validate` runs all five complete observations once, without collecting timings.

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.CSV.Benchmarks.Modern -- --filter "*MarkPflug65KModernCsvBenchmarks*" --inProcess --warmupCount 8 --iterationCount 12 --invocationCount 8 --unrollFactor 1 --outliers DontRemove
```

I retain the complete BenchmarkDotNet JSON, environment, package and source
identities, and all samples for a comparison. Different frameworks and machines
are separate cohorts. `OFFICEIMO_BENCHMARK_PROCESS_PRIORITY=Normal` selects the
same optional process-priority control as the other runners. CsvHelper is used under its Apache-2.0 license option;
Sep and BenchmarkDotNet are MIT. FourLambda is MPL-2.0 and uses its software
fallback on hosts without AVX2; I record that platform distinction with the results.
All comparison dependencies remain local to
this nonpackable executable.
