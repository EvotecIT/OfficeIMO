# OfficeIMO.Data.Conversions.Benchmarks

Opt-in comparison of structured-data conversions against the pipelines people build from established libraries: YamlDotNet's JSON-compatible serializer, Newtonsoft.Json's XML conversion, and CsvHelper with System.Text.Json. These libraries are referenced only here.

```powershell
dotnet run -c Release --project OfficeIMO.Data.Conversions.Benchmarks -- --filter '*' --job short
```

Setup validates OfficeIMO's output for every lane (document, record and row counts) before timing. The CSV-to-JSON comparison lane writes strings only, while OfficeIMO infers typed columns, so it does less work than OfficeIMO.
