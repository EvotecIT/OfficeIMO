# OfficeIMO.Adf benchmarks

This opt-in project measures complete Atlas Document Format JSON parsing and
semantic parse/write round trips at small and normal document scales. Every
lane validates the entire JSON tree, including unknown root, node, and mark
properties.

The `System.Text.Json` lane is a cost floor, not a feature-equivalent
competitor: it creates a mutable JSON tree but does not create OfficeIMO's typed
ADF model or perform ADF validation. Ratios therefore quantify the incremental
cost of the typed contract and must not be described as product parity.

Run BenchmarkDotNet on the selected runtime:

```powershell
dotnet run --project OfficeIMO.Adf.Benchmarks/OfficeIMO.Adf.Benchmarks.csproj -c Release -f net8.0 -- --job short
```

Capture isolated time, allocation, retained-heap, managed-peak, process-peak,
and input/output-size evidence:

```powershell
dotnet run --project OfficeIMO.Adf.Benchmarks/OfficeIMO.Adf.Benchmarks.csproj -c Release -f net8.0 -- evidence --repeat 3 --json .benchmark-artifacts/adf/evidence.json
```

Generated benchmark artifacts are machine-specific and remain outside source
control.
