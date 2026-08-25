# OfficeIMO.GoogleWorkspace.Benchmarks

This opt-in project measures the public dependency-light Google Workspace transport without network latency. The deterministic handler returns the same declared-length payload on every request. The OfficeIMO lane still performs its public retry, timeout, safety, response-limit, diagnostics, and response-copy path; raw `HttpClient` is not treated as equivalent work because it omits those contracts.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.GoogleWorkspace.Benchmarks -- --filter '*GoogleWorkspaceTransportBenchmarks*'
```

BenchmarkDotNet validates the returned bytes during setup and measures public `SendBytesAsync(...)` time and managed allocation at 64 KiB and 4 MiB. Raw artifacts belong under `.benchmark-artifacts` and should not be committed.

Generate isolated retained/peak-memory, exact size/hash, and source-provenance evidence:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.GoogleWorkspace.Benchmarks -- evidence --json .benchmark-artifacts\googleworkspace-transport\evidence.json
```
