# OfficeIMO.Confluence managed-section performance

This is a Windows baseline for the pure `ConfluenceManagedSection.Apply` contract. It measures page scanning, marker validation, exact replacement, the complete updated body, and both public UTF-8 SHA-256 hashes. Network transport and authentication are deliberately outside the lane.

## Environment and provenance

- AMD Ryzen 9 9950X3D2, 32 logical / 16 physical cores
- Windows 11 25H2, build 26200.9168
- .NET SDK 10.0.111
- .NET 8.0.30 x64 runtime
- BenchmarkDotNet 0.15.8 `ShortRun`: one launch, three warmups, three measured iterations
- optimized source commit `13a0edec2e715e4c7d0c1eca8042e1dd8b4b7d3c`, clean worktree for isolated evidence

Commands:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Confluence.Benchmarks -- --filter '*ConfluenceManagedSectionBenchmarks*'
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Confluence.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\confluence\evidence.json
```

## Improvement

The initial implementation allocated complete UTF-8 byte arrays for both hashes and materialized prefix and suffix substrings before building the updated body. Hashing now streams bounded UTF-8 chunks through a cleared shared buffer, including correct surrogate-pair handling across chunk boundaries. Modern .NET fills the one required result string directly.

| Page scale | Before mean / allocation | After mean / allocation | Change |
| --- | ---: | ---: | ---: |
| 16 KiB | 22.70 μs / 103.88 KiB | 19.80 μs / 38.02 KiB | 12.8% less time / 63.4% less allocation |
| 1 MiB | 1.546 ms / 6.19 MiB | 1.166 ms / 2.13 MiB | 24.6% less time / 65.6% less allocation |

The 1 MiB allocation is now approximately the required UTF-16 updated result plus small fixed bookkeeping. `netstandard2.0` and .NET Framework use the compatible substring construction path but still receive bounded pooled hashing; all 22 contract tests pass on .NET Framework 4.7.2, .NET 8, and .NET 10.

## Isolated memory and output evidence

Each evidence sample runs in a fresh child process after warmup. Hashes and output sizes were identical across all three repetitions.

| Page scale | Median elapsed | Allocated | Retained | Managed peak growth | Process working-set growth | Input / output bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 16 KiB | 0.0264 ms | 39,184 | 37,208 | 60,464 | 294,912 | 16,603 / 17,631 |
| 1 MiB | 1.3043 ms | 2,232,816 | 2,230,504 | 2,261,352 | 2,383,872 | 1,048,711 / 1,114,281 |

The evidence runner also records absolute process peak, source commit, dirty-tree state, runtime, operating system, architecture, and semantic hashes in its local JSON output.

## Comparison boundary

No competitor ratio is accepted for this narrow Confluence ownership contract. The implementation already uses .NET ordinal search and SHA-256 primitives; constructing a benchmark-only duplicate with the same OfficeIMO markers, validation rules, exact body, and two hashes would be a synthetic baseline rather than independent competition. A package comparison should be added only if another public library exposes equivalent managed-section behavior.

The remaining work is cross-platform evidence and a lower-copy compatible construction path for `netstandard2.0` and .NET Framework. The current result is a substantial internal improvement, not a claim of optimality.
