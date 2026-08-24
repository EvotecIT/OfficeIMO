# OfficeIMO.Visio package-workflow evidence (2026-08-24)

## Result

OfficeIMO.Visio now has reproducible owner-level performance coverage for
complete VSDX creation/save and load/structural-inspection workflows. Reusing
page-local automatic-ID bookkeeping and streaming the completed package into
the caller's destination removed avoidable copies and temporary collections.
The latest pass also replaces repeated whole-page identifier scans with a
page-local index and avoids boxed child-shape enumerators in internal tree
traversal. Against the original source baseline, managed allocation for
creation is now 16.0% lower on the small corpus, 44.1% lower on the normal
corpus, and 64.4% lower on the large corpus. Large load/inspection allocation
is 31.0% lower.

This is an owner baseline, not a contender-ratio claim. Aspose.Diagram is a
possible commercial comparison, but its documented
[evaluation mode](https://docs.aspose.com/diagram/net/licensing/) limits loaded
diagrams and watermarks saved output. A public ratio requires a separately
licensed, opt-in runner whose output passes the same structural contract.
Microsoft Visio automation is also not an equivalent managed, cross-platform
library workflow.

## Validated contract

The deterministic matrix covers:

| Scale | Pages | Shapes | Connectors | Shape Data rows | VSDX bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| Small | 1 | 25 | 24 | 25 | 47,522 |
| Normal | 4 | 400 | 396 | 400 | 678,667 |
| Large | 8 | 2,000 | 1,992 | 2,000 | 3,350,485 |

Every generated package is reopened before measurement is accepted. Validation
checks page, shape, connector, and Shape Data counts plus boundary text on the
first and last shapes. The load lane materializes a public structural snapshot
from the same validated package. This prevents a faster but incomplete writer
or reader from being treated as an improvement.

The production changes preserve automatic-ID collision handling, including
nested group children, child mutations after a group is attached to a page,
released-ID reuse, and IDs reserved during relationship construction. Group
trees retain duplicate-ID validation. Saving still rewinds and truncates
caller-provided seekable streams; it copies the completed package directly
instead of first materializing a second full byte array.

## BenchmarkDotNet results

Environment: Windows 11 build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 `ShortRun`. The final source was clean
at commit `9f0f999213418cb7498dc4137b7c7a9983c485a6`.

| Operation | Scale | Mean | Managed allocation |
| --- | --- | ---: | ---: |
| Create and save | Small | 0.963 ms | 828.90 KiB |
| Create and save | Normal | 8.931 ms | 8,911.55 KiB |
| Create and save | Large | 68.685 ms | 48,019.15 KiB |
| Load and inspect | Small | 2.365 ms | 1,450.16 KiB |
| Load and inspect | Normal | 15.568 ms | 13,705.58 KiB |
| Load and inspect | Large | 105.928 ms | 65,307.29 KiB |

The same-machine initial production source was commit
`62f01ceb154af3c3194b06273585d70b0f3428a4`, with the benchmark harness applied
only to measure it. Allocation changed as follows:

| Workload | Initial | Final | Change |
| --- | ---: | ---: | ---: |
| Create and save, Small | 987.00 KiB | 828.90 KiB | 16.0% lower |
| Create and save, Normal | 15,943.63 KiB | 8,911.55 KiB | 44.1% lower |
| Create and save, Large | 134,928.59 KiB | 48,019.15 KiB | 64.4% lower |
| Load and inspect, Small | 1,473.17 KiB | 1,450.16 KiB | 1.6% lower |
| Load and inspect, Normal | 15,992.46 KiB | 13,705.58 KiB | 14.3% lower |
| Load and inspect, Large | 94,678.68 KiB | 65,307.29 KiB | 31.0% lower |

The short timing run remains noisy, so the result is not a blanket speed claim.
The same-machine large means improved by 29.2% for creation and 17.6% for
load/inspection, while the normal load mean varied upward. Allocation is the
more repeatable improvement; elapsed-time stability and remaining XML,
preservation, inspection-graph, and package allocation stay open targets.

## Isolated memory and output evidence

The evidence runner starts a fresh child process for every scale, operation,
and repetition. Values below are medians over three repetitions. Managed peak
is sampled over the measurement batch; process peak is the absolute child
working-set peak. The evidence identifies the clean source commit and runtime.

| Operation | Scale | Time/op | Allocation/op | Retained heap | Managed batch peak | Process peak | Package bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create and save | Small | 1.19 ms | 0.88 MiB | 47.4 KiB | 28.12 MiB | 72.07 MiB | 47,522 |
| Create and save | Normal | 15.77 ms | 9.74 MiB | 663.5 KiB | 11.93 MiB | 64.06 MiB | 678,667 |
| Create and save | Large | 89.66 ms | 52.20 MiB | 3.20 MiB | 60.64 MiB | 132.36 MiB | 3,350,485 |
| Load and inspect | Small | 1.97 ms | 1.42 MiB | 46.4 KiB | 45.52 MiB | 90.95 MiB | 47,522 |
| Load and inspect | Normal | 36.72 ms | 13.63 MiB | 728.9 KiB | 47.65 MiB | 110.74 MiB | 678,667 |
| Load and inspect | Large | 165.24 ms | 65.03 MiB | 3.55 MiB | 77.17 MiB | 142.09 MiB | 3,350,485 |

The isolated runner includes validation and process startup behavior outside its
timed loop, while BenchmarkDotNet supplies the primary steady-state time and
allocation measurements. Large creation now allocates roughly 16 times the
resulting package size, and large loading roughly 20 times the input size.
Those ratios are owner-level optimization signals, not comparisons with another
implementation.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- --filter '*VisioBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\visio\evidence.json
```

Raw BenchmarkDotNet and process evidence remain ignored machine-local
artifacts. This note retains the compact reproducible result and exact source
provenance.
