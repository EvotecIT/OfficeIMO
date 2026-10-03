# OfficeIMO LaTeX benchmarks

This opt-in BenchmarkDotNet project measures complete lossless parsing and
parse-plus-preserve-write and complete public-token inspection workflows over deterministic small, normal, and large
documents. Every workload is validated before timing: parsing must be lossless,
diagnostics must contain no errors, heading counts and content markers must
match, and preserve writing must reproduce the input exactly.

There is currently no general .NET comparison lane. The available libraries
found during the 2026-08-24 inventory either expose only a much smaller LaTeX
subset or do not provide a public parse entry point that can execute this
workflow. Add a comparison only when both implementations perform equivalent
work and both outputs can be validated.

Validate semantics and input/output size:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Latex.Benchmarks -- validate
```

Capture isolated elapsed, allocation, retained-heap, managed-heap peak, process
peak, and output-size evidence, or enforce the checked-in regression budgets:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Latex.Benchmarks -- --evidence --repeat 3 --json .benchmark-artifacts\latex\evidence.json
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Latex.Benchmarks -- --verify-budgets
```

Measure full public-token inspection separately:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Latex.Benchmarks -- --evidence --operation ParseInspect --repeat 3 --json .benchmark-artifacts\latex\inspection.json
```

`ParseInspect` reads and validates every token's text, value, span, and termination
state, and retains the parsed document. This exposes allocation deferred by compact
internal token storage. It has no checked-in budget; the default evidence and budget
routes continue to select `Parse` and `ParseWrite`. Each isolated measurement records
the SHA-256 hashes of the loaded native and benchmark assemblies. BenchmarkDotNet
setup reports the same hashes and validates the complete inspection path.

The corpus uses fixed CRLF line endings on every OS; input hashes identify the
exact source measured. Retained memory is live heap size minus fragmentation after
forced full blocking collection. The report keeps the older `GC.GetTotalMemory`
retained estimate separately and labels sampled peaks as estimates. Process peak
is `null` when the platform API does not report it; budget verification then fails
with an unavailable-metric diagnostic. A missing value is never interpreted as zero.

Start with a dry run, then use the short or default job for measurements:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Latex.Benchmarks -- --filter '*LatexParseBenchmarks*' --job Dry --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Latex.Benchmarks -- --filter '*LatexParseBenchmarks*' --job Short --noOverwrite
```

BenchmarkDotNet artifacts are machine-specific. Keep them under the ignored
`.benchmark-artifacts` root or another temporary location and publish only
environment-qualified results.
