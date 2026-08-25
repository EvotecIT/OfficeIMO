# OfficeIMO.Confluence benchmarks

This opt-in project measures pure, deterministic Confluence workloads without network or authentication noise. The managed-section lane validates the complete updated body and both public SHA-256 hashes before BenchmarkDotNet measures it.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Confluence.Benchmarks -- --filter '*ConfluenceManagedSectionBenchmarks*'
```

The 16 KiB page represents ordinary managed content. The 1 MiB page exposes file-sized scans, replacement construction, and hashing allocation. Benchmark artifacts are machine-specific and stay outside the repository.

Generate isolated allocation, retained-heap, managed-peak, process-peak, and input/output-size evidence:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Confluence.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\confluence\evidence.json
```
