# OfficeIMO.GoogleWorkspace transport performance

This Windows baseline measures the public dependency-light `GoogleWorkspaceHttpTransport.SendBytesAsync(...)` path with deterministic in-memory HTTP content. It excludes network latency while retaining OfficeIMO request safety, retry, timeout, diagnostic, and maximum-response-size handling.

## Environment and provenance

- AMD Ryzen 9 9950X3D2, 32 logical / 16 physical cores
- Windows 11 25H2, build 26200.9168
- .NET SDK 10.0.111
- .NET 8.0.30 x64 runtime
- BenchmarkDotNet 0.15.8 `ShortRun`: one launch, three warmups, three measured iterations
- optimized source commit `1d2096ae2dee248554d211a2f5ab37c1bfae13d2`
- clean worktree for isolated evidence

Commands:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.GoogleWorkspace.Benchmarks -- --filter '*GoogleWorkspaceTransportBenchmarks*'
dotnet run -c Release -f net8.0 --project .\OfficeIMO.GoogleWorkspace.Benchmarks -- evidence --json .benchmark-artifacts\googleworkspace-transport\evidence.json
```

## Declared-length response improvement

The previous bounded reader copied a declared-length response into a growable `MemoryStream`, allocated a fixed 80 KiB scratch array, and then copied the complete stream into the returned array. The optimized path validates `Content-Length` against the configured limit, allocates the required return array once, fills it directly, and still probes for dishonest extra bytes. Unknown-length content retains bounded streaming while using pooled scratch storage.

| Payload | Before mean / allocation | After mean / allocation | Change |
| ---: | ---: | ---: | ---: |
| 64 KiB | 6.486 us / 211.66 KiB | 3.113 us / 67.63 KiB | 52.0% less time / 68.1% less allocation |
| 4 MiB | 2.268 ms / 14,340.09 KiB | 0.476 ms / 4,099.88 KiB | 4.77x faster / 71.4% less allocation |

The 4 MiB result now allocates only about 3.9 KiB beyond the required returned byte array. The benchmark setup rejects any result that differs from the deterministic input payload.

## Isolated memory and size evidence

Each payload runs in a fresh child process after two warmups. The clean evidence records the exact input/output byte count and SHA-256 hash as well as retained and sampled peak memory.

| Payload | Allocated | Retained | Managed peak growth | Process working-set growth | Absolute process peak | Output |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 64 KiB | 79,336 | 72,720 | 119,768 | 417,792 | 32,923,648 | 65,536 bytes, hash validated |
| 4 MiB | 4,218,784 | 4,212,144 | 4,252,272 | 839,680 | 45,559,808 | 4,194,304 bytes, hash validated |

The retained heap intentionally includes the returned payload. Working-set sampling at these sub-millisecond durations is diagnostic rather than a timing gate; BenchmarkDotNet owns the elapsed/allocation claim.

## Comparison boundary

Raw `HttpClient.ReadAsByteArrayAsync` is not accepted as an equivalent competitor because it omits OfficeIMO retry policy, per-request deadlines, safety classification, diagnostics, response limits, and mutation outcome handling. This is a before/after owner improvement, not a competitor ratio. A future external lane must preserve those contracts before its numbers become contender evidence.
