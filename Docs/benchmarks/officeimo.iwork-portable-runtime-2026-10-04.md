# iWork portable runtime evidence — 2026-10-04

The [opt-in workflow run](https://github.com/EvotecIT/OfficeIMO/actions/runs/37231067068) records 6,390 validated samples across Linux x64, macOS arm64 and Windows x64. Each host completes 41 fresh PowerShell runs and 2,130 samples. Raw JSON/CSV counters agree, semantic and active I/O cancellation checks pass, repeated managed baselines remain fixed within each case/operation, and measured assembly/spec hashes remain consistent within each host. Slower observations remain in the results.

The OfficeIMO workload source is [`e7e5993a5a8eab43b83e85b3f886731a587083db`](https://github.com/EvotecIT/OfficeIMO/commit/e7e5993a5a8eab43b83e85b3f886731a587083db); the canonical PowerForge owner is [`000237c50bad900df3ce95b7ed81488fa1189734`](https://github.com/EvotecIT/PSPublishModule/commit/000237c50bad900df3ce95b7ed81488fa1189734). The workload assemblies target .NET 8. Measurements execute in the actual PowerShell runtime below; they are not .NET 8 timing measurements.

| Host | Architecture | Logical CPUs | PowerShell | Measurement runtime |
| --- | --- | ---: | --- | --- |
| Linux | X64 | 4 | 7.6.6 | .NET 10.0.12 |
| macOS | Arm64 | 5 | 7.6.5 | .NET 10.0.11 |
| Windows | X64 | 4 | 7.6.6 | .NET 10.0.12 |

Each host runs the synthetic scale matrix twice with five measured iterations per case/operation. Native Pages, Numbers and Keynote load, conversion/save and package-read cancellation run 50 times per case/operation in each of two passes. Native Keynote copying has the same repeated cancellation contract. A separate 50-iteration native pass enables 5 ms managed/resident sampling. Warmups, input preparation, full readback and explicit collections stay outside operation timing. The [workload contract](../../OfficeIMO.IWork.Benchmarks/README.md) defines the fixture hashes, policy, sizes and cancellation boundaries.

## Uninstrumented conversion

Values below retain both passes as a range of their median milliseconds. These are observations on the recorded hosted machines, with host invocation included. The largest synthetic cases contain 10,000 Pages paragraphs, 100,000 Numbers cells and 1,000 Keynote slides. Native fixtures are small independent-producer packages.

| Case | Linux median ms | macOS median ms | Windows median ms |
| --- | ---: | ---: | ---: |
| Pages-Large | 102.73–108.13 | 53.75–62.24 | 89.80–90.40 |
| Numbers-Large | 271.61–290.75 | 246.13–276.24 | 238.30–253.34 |
| Keynote-Large | 669.19–678.72 | 526.05–551.00 | 596.82–610.98 |
| Pages-Native | 39.73–40.89 | 25.89–26.22 | 27.78–33.55 |
| Numbers-Native | 44.54–46.58 | 25.08–26.00 | 31.66–32.14 |
| Keynote-Native | 62.38–63.23 | 36.01–37.23 | 46.35–47.39 |

Managed allocation includes process-wide host and concurrent-thread allocation. Mean allocations for the large synthetic conversions are approximately 48–49 MiB for Pages, 253–254 MiB for Numbers and 516–517 MiB for Keynote. They are allocation volume, not retained or peak usage.

## Operation memory observations

The table gives the largest sampled maximum-minus-baseline change among 50 instrumented native conversion iterations. Sampling and normal timing use separate run modes. Sampled maxima are lower bounds on peaks; short operations may have only boundary readings. These values include observer/host effects and do not isolate native allocations.

| Host | Native family | Observed managed change MiB | Observed resident change MiB |
| --- | --- | ---: | ---: |
| Linux | Pages | 14.67 | 0.51 |
| Linux | Numbers | 19.09 | 11.59 |
| Linux | Keynote | 22.82 | 3.98 |
| macOS | Pages | 9.54 | 3.78 |
| macOS | Numbers | 8.59 | 5.00 |
| macOS | Keynote | 13.31 | 6.69 |
| Windows | Pages | 13.24 | 4.08 |
| Windows | Numbers | 17.35 | 7.99 |
| Windows | Keynote | 22.21 | 11.47 |

For the uninstrumented native conversions, the final collected managed heap relative to the first measured setup baseline is 2.71–2.91 MiB on Linux, 2.79–2.90 MiB on Windows and 5.58–9.17 MiB on macOS. These fixed-baseline observations include retained benchmark samples, PowerShell state and caches. They do not diagnose an OfficeIMO leak or native-memory retention. Positive changes require controlled attribution before choosing a leak threshold.

## Active I/O cancellation

All 3,150 cancellation samples request cancellation after actual package reads or native stream writes, observe the exception and preserve caller stream ownership. Native copying stops at the first 64 KiB. The worst recorded request-to-exception latency includes both normal and instrumented passes:

| Host | Maximum observed request-to-exception ms |
| --- | ---: |
| Linux | 0.1103 |
| macOS | 0.1167 |
| Windows | 0.1425 |

These fixture-specific latencies are observations, not portable deadlines. They qualify package intake and native stream copying. Projection, destination construction, native encoding and atomic path staging still require mid-operation latency evidence.

## Reproduction and budget boundary

Dispatch `.github/workflows/iwork-runtime-evidence.yml` at the recorded workload source, or run `Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1` in fresh PowerShell hosts using the pinned source-built module. Keep the artifact JSON, CSV, summaries and environment metadata together. Use the canonical summary for a host-specific baseline, retain failed and slower samples, and compare equivalent runtime, workload, policy and instrumentation.

The two passes establish reproducible observations on three available hosted targets. They do not establish universal elapsed/allocation ceilings, exact managed/resident peaks, native-memory retention, large independent-producer budgets or a portable leak threshold. Measurements stay opt-in. Those contracts remain in [I5](../ROADMAP.md#i5-runtime-and-apple-host-acceptance).
