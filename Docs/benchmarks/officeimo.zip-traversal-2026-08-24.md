# OfficeIMO.Zip safe traversal evidence - 2026-08-24

This run measures the overhead of `OfficeIMO.Zip` safety policy over direct
`System.IO.Compression` metadata traversal. Both lanes open the same in-memory
ZIP, sort entries ordinally, exclude directory entries, and materialize the
same path, name, depth, length, and timestamp fields.

The platform lane does not enforce OfficeIMO's path-traversal, depth,
entry-count, total-size, per-entry-size, or expansion-ratio limits. The ratios
below are the cost of those additional checks on a safe corpus, not evidence
that the platform lane has the same defensive contract.

## Environment and validation

- Source commit: `6bd01a7c0a00212af95a8c832e6bc2c9c4d6f573`
- Source tree: clean
- Runtime: .NET 10.0.11, x64 RyuJIT
- SDK: 10.0.111
- OS: Windows 11 25H2, build 26200.9168
- CPU: AMD Ryzen 9 9950X3D2, 32 logical cores
- BenchmarkDotNet: 0.15.8, full default job

Preflight validation requires exact agreement on ordered paths, names,
directory flags, depths, uncompressed lengths, UTC timestamps, total bytes, and
a SHA-256 structural fingerprint. The inputs contain 24, 512, and 4,000 files
in reverse package order, plus excluded directory entries. Their ZIP sizes are
6,288, 163,061, and 1,525,887 bytes.

## Full timing and allocation result

| Scale | OfficeIMO mean | Platform mean | Time ratio | OfficeIMO median | Platform median | OfficeIMO allocated | Platform allocated | Allocation ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Small | 4.959 us | 4.466 us | 1.11x | 4.902 us | 4.388 us | 27.37 KB | 27.21 KB | 1.01x |
| Normal | 133.382 us | 114.687 us | 1.16x | 123.705 us | 117.861 us | 466.56 KB | 466.41 KB | 1.00x |
| Large | 2,863.426 us | 2,768.222 us | 1.03x | 2,864.870 us | 2,760.639 us | 3,527.54 KB | 3,527.39 KB | 1.00x |

The normal workload was multimodal for both implementations (`mValue` 3.30
and 3.94). Mean and median are both shown so the host variance remains visible.
Even the more conservative mean ratio shows modest safety-policy overhead on
this benign corpus. Because the raw platform lane does not enforce that policy,
this evidence is diagnostic and is not published as a contender comparison.

The optimization replaced per-entry path splitting with a zero-allocation
segment scan, removed a duplicate options normalization on stream traversal,
and pre-sized the accepted-entry list. Compared with the preceding short-run
baseline, OfficeIMO allocation fell by 14.9-17.6 percent and now matches the
platform projection at all three scales.

## Isolated peak managed heap

Five fresh child processes per engine and scale recorded input size, total
allocation, and sampled peak managed-heap growth. Timing from these probes
includes cold JIT and is not used for the contender ratios above.

| Scale | OfficeIMO peak growth | Platform peak growth | OfficeIMO allocated | Platform allocated |
| --- | ---: | ---: | ---: | ---: |
| Small | 68,072 B | 67,944 B | 51,000 B | 48,544 B |
| Normal | 518,856 B | 518,728 B | 500,736 B | 498,280 B |
| Large | 3,660,600 B | 3,660,472 B | 3,635,120 B | 3,632,664 B |

Peak managed-heap growth is effectively identical. There is no output-size
comparison: `OfficeIMO.Zip` owns traversal policy, not ZIP creation, and both
lanes consume the same validated input package.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Zip.Benchmarks -- validate
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload ziptraverse -RunMode full -Framework net10.0
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Zip.Benchmarks -- --evidence --repeat 5 --json .benchmark-artifacts\zip\evidence.json
```

Raw BenchmarkDotNet output and isolated probe JSON remain local, while this
small summary records the reproducible contract, environment, and result.
