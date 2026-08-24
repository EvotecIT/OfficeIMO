# OfficeIMO.Adf benchmarks

This opt-in project measures complete Atlas Document Format JSON parsing and
semantic parse/write round trips at small and normal document scales. Every
lane validates the entire JSON tree, including unknown root, node, and mark
properties.

The `System.Text.Json` tree lane is the narrowest cost floor. The typed-model
lane additionally creates a typed document/node/mark graph and preserves
extension data, making it the primary baseline. It still does not perform ADF
structural validation and is not a full feature-equivalent competitor. Ratios
therefore quantify implementation overhead and must not be described as product
parity.

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
