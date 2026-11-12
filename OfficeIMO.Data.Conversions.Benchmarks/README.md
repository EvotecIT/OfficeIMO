# OfficeIMO.Data.Conversions.Benchmarks

Opt-in comparison of structured-data conversions against the pipelines people build from established libraries: YamlDotNet's JSON-compatible serializer, Newtonsoft.Json's XML conversion, CsvHelper with System.Text.Json (comma and tab delimiters), and Tomlyn's model with System.Text.Json. These libraries are referenced only here.

```powershell
dotnet run -c Release --project OfficeIMO.Data.Conversions.Benchmarks -- --filter '*' --job short
```

Setup validates OfficeIMO's output for every lane (document, record and row counts) before timing. The CSV-to-JSON and TSV-to-JSON comparison lanes write strings only, while OfficeIMO infers typed columns, so they do less work than OfficeIMO. The TOML lanes use OfficeIMO's mapping on both sides: date-times as RFC 3339 strings, nulls omitted and a root array wrapped in `root`.

`--verify` has the comparison libraries read OfficeIMO's output, including Tomlyn reading TOML written from the JSONTestSuite accept cases. `--profile` times single lanes with a stopwatch for quick iteration.
