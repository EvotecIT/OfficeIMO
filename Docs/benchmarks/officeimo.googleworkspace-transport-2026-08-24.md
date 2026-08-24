# OfficeIMO.GoogleWorkspace transport performance

This Windows baseline measures the public dependency-light `GoogleWorkspaceHttpTransport.SendBytesAsync(...)` path with deterministic in-memory HTTP content. It excludes network latency while retaining OfficeIMO request safety, retry, timeout, diagnostic, and maximum-response-size handling.

## Environment and provenance

- AMD Ryzen 9 9950X3D2, 32 logical / 16 physical cores
- Windows 11 25H2, build 26200.9168
- .NET SDK 10.0.111
- .NET 8.0.30 x64 runtime
- BenchmarkDotNet 0.15.8 `ShortRun`: one launch, three warmups, three measured iterations
- optimized source commit `ab919b5a9d6ffb29a4d19e14d5a88fc9ce32b298`
- clean worktree for isolated evidence

Commands:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.GoogleWorkspace.Benchmarks -- --filter '*GoogleWorkspaceTransportBenchmarks*' --artifacts .benchmark-artifacts\googleworkspace-transport
dotnet run -c Release -f net8.0 --project .\OfficeIMO.GoogleWorkspace.Benchmarks -- evidence --json .benchmark-artifacts\googleworkspace-transport\evidence.json
```

## Declared-length response improvement

The previous bounded reader copied a declared-length response into a growable `MemoryStream`, allocated a fixed 80 KiB scratch array, and then copied the complete stream into the returned array. The optimized path validates `Content-Length` against the configured limit, allocates the required return array once, fills it directly, and still probes for dishonest extra bytes.

| Payload | Before mean / allocation | After mean / allocation | Change |
| ---: | ---: | ---: | ---: |
| 64 KiB | 6.486 us / 211.66 KiB | 3.113 us / 67.63 KiB | 52.0% less time / 68.1% less allocation |
| 4 MiB | 2.268 ms / 14,340.09 KiB | 0.476 ms / 4,099.88 KiB | 4.77x faster / 71.4% less allocation |

The 4 MiB result now allocates only about 3.9 KiB beyond the required returned byte array. The benchmark setup rejects any result that differs from the deterministic input payload.

## Unknown-length contender gap

The first unknown-length implementation still copied a growing `MemoryStream` into the returned array. That left the 4 MiB path at 3.75x declared-length time and 2.97x allocation, which is a material deficit under the repository's 2x contender boundary. The current implementation streams into cleared pooled chunks and performs one exact final allocation.

| Payload | Declared mean / allocation | Unknown mean / allocation | Unknown / declared | Previous ratio |
| ---: | ---: | ---: | ---: | ---: |
| 64 KiB | 3.866 us / 67.63 KiB | 6.035 us / 67.59 KiB | 1.56x time / 1.00x allocation | 1.72x / 1.94x |
| 4 MiB | 558.692 us / 4,100.09 KiB | 723.621 us / 4,100.65 KiB | 1.30x time / 1.00x allocation | 3.75x / 2.97x |

Both current lanes are inside 2x, but this is contender-level rather than an optimality claim. The paired ratios matter more than differences between separate short runs, especially for sub-millisecond timing.

## Isolated memory and size evidence

Each payload runs in a fresh child process after two warmups. The clean evidence records the exact input/output byte count and SHA-256 hash as well as retained and sampled peak memory.

| Payload | Length | Allocated | Retained | Managed peak growth | Process working-set growth | Absolute process peak | Output |
| ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 64 KiB | Declared | 79,416 | 72,720 | 119,768 | 413,696 | 33,271,808 | 65,536 bytes, hash validated |
| 64 KiB | Unknown | 79,112 | 72,720 | 119,808 | 475,136 | 33,579,008 | 65,536 bytes, same hash |
| 4 MiB | Declared | 4,208,304 | 4,205,600 | 4,244,104 | 495,616 | 45,604,864 | 4,194,304 bytes, hash validated |
| 4 MiB | Unknown | 4,209,240 | 4,205,600 | 4,244,104 | 491,520 | 50,163,712 | 4,194,304 bytes, same hash |

The retained heap intentionally includes the returned payload. Working-set sampling at these sub-millisecond durations is diagnostic rather than a timing gate; BenchmarkDotNet owns the elapsed/allocation claim.

## Comparison boundary

Raw `HttpClient.ReadAsByteArrayAsync` is not accepted as an equivalent competitor because it omits OfficeIMO retry policy, per-request deadlines, safety classification, diagnostics, response limits, and mutation outcome handling. This is a before/after owner improvement, not a competitor ratio. A future external lane must preserve those contracts before its numbers become contender evidence.
