# OfficeIMO AsciiDoc benchmarks

This opt-in BenchmarkDotNet project measures the complete source-preserving
parse and parse-plus-preserve-write workflows at three deterministic scales.
Every workload validates lossless parsing, headings, tables, semantic markers,
and byte-identical preserve output before timing begins.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.AsciiDoc.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.AsciiDoc.Benchmarks -- --filter '*AsciiDocParseBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.AsciiDoc.Benchmarks -- --evidence --repeat 3 --json .\.benchmark-artifacts\asciidoc\evidence.json
dotnet run -c Release -f net10.0 --project .\OfficeIMO.AsciiDoc.Benchmarks -- --verify-budgets
```

The isolated evidence lane starts a fresh child process for each measurement and
records elapsed time, total managed allocation, retained managed heap, sampled
managed-heap peak, absolute process peak working set, and UTF-8 input/output
bytes. The JSON budgets are regression ceilings for this machine-independent
contract, not a claim that the current costs are optimal or competitively
acceptable. Preserve-write output must remain byte-identical to the input.

AsciiDocNet is not included because its documented unsupported tables and list
continuations do not satisfy this workload. Keep any future third-party package
isolated to an opt-in comparison project and require equivalent validated work.
