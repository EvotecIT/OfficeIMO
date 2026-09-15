# OfficeIMO provenance benchmarks

This opt-in project measures bounded structural C2PA carrier inspection and
selective removal across PNG, TIFF, SVG, ZIP packages, and structured text.
Each deterministic fixture contains one structurally valid manifest store.
Validation requires exact format detection, one valid carrier, one removal,
and no remaining carrier. PNG, TIFF, SVG, and text output must match exact
expected bytes; ZIP output must contain exactly the preserved payload entry
with byte-identical content.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- --filter '*ProvenanceBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Provenance.Benchmarks -- evidence --repeat 3 --json .\.benchmark-artifacts\provenance\evidence.json --budget .\Build\BenchmarkBudgets\security-provenance.json
```

The `evidence` lane runs every case in a fresh child process and records elapsed
time, managed allocations, retained and peak managed heap growth, absolute
process peak working set, input size, and exact removal-output size. The JSON
also records the source commit and whether the source tree was dirty.
The checked-in large-fixture budgets cover every format on Windows, Linux, and
macOS. They gate elapsed time, allocation, managed peak, process peak, and
removal-output size while the preflight keeps content fidelity authoritative.

The project stays outside `OfficeIMO.sln`. No external library is presented as
a performance comparator because no official managed .NET C2PA library exposes
the same bounded structural inspect-and-selective-remove contract. The existing
`OfficeIMO.Provenance.C2pa` adapter and host-supplied `c2patool` remain the
cryptographic interoperability and verification lane.
