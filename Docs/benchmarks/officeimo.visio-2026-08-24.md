# OfficeIMO.Visio package-workflow evidence (2026-08-24)

## Result

OfficeIMO.Visio now has reproducible owner-level performance coverage for
complete VSDX creation/save and load/structural-inspection workflows. Reusing
page-local automatic-ID bookkeeping and streaming the completed package into
the caller's destination removed avoidable copies and temporary collections.
Managed allocation for creation fell by 11.1% on the small corpus, 24.8% on the
normal corpus, and 35.7% on the large corpus.

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
nested group children and IDs reserved during relationship construction. The
leaf-shape path avoids a temporary set, while group trees retain duplicate-ID
validation. Saving still rewinds and truncates caller-provided seekable streams;
it now copies the completed package directly instead of first materializing a
second full byte array.

## BenchmarkDotNet results

Environment: Windows 11 build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 `ShortRun`. The final source was clean
at commit `0352cce4b011bf05deab8cd433293d95deecdc1c`.

| Operation | Scale | Mean | Managed allocation |
| --- | --- | ---: | ---: |
| Create and save | Small | 2.121 ms | 877.58 KiB |
| Create and save | Normal | 25.591 ms | 11,983.29 KiB |
| Create and save | Large | 97.068 ms | 86,796.99 KiB |
| Load and inspect | Small | 1.936 ms | 1,467.96 KiB |
| Load and inspect | Normal | 11.779 ms | 15,908.36 KiB |
| Load and inspect | Large | 128.620 ms | 93,987.08 KiB |

The same-machine initial production source was commit
`62f01ceb154af3c3194b06273585d70b0f3428a4`, with the benchmark harness applied
only to measure it. Allocation changed as follows:

| Workload | Initial | Final | Change |
| --- | ---: | ---: | ---: |
| Create and save, Small | 987.00 KiB | 877.58 KiB | 11.1% lower |
| Create and save, Normal | 15,943.63 KiB | 11,983.29 KiB | 24.8% lower |
| Create and save, Large | 134,928.59 KiB | 86,796.99 KiB | 35.7% lower |
| Load and inspect, Small | 1,473.17 KiB | 1,467.96 KiB | 0.4% lower |
| Load and inspect, Normal | 15,992.46 KiB | 15,908.36 KiB | 0.5% lower |
| Load and inspect, Large | 94,678.68 KiB | 93,987.08 KiB | 0.7% lower |

The short timing run is not stable enough for a blanket speed claim. Normal
creation and large loading were faster in this run, while small and large
creation and normal loading varied in both directions across runs. Allocation
is the repeatable improvement; elapsed-time and remaining graph/package
allocation stay open optimization targets.

## Isolated memory and output evidence

The evidence runner starts a fresh child process for every scale, operation,
and repetition. Values below are medians over three repetitions. Managed peak
is sampled over the measurement batch; process peak is the absolute child
working-set peak. The evidence identifies the clean source commit and runtime.

| Operation | Scale | Time/op | Allocation/op | Retained heap | Managed batch peak | Process peak | Package bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create and save | Small | 1.64 ms | 0.92 MiB | 47.4 KiB | 29.54 MiB | 73.63 MiB | 47,522 |
| Create and save | Normal | 23.59 ms | 12.73 MiB | 663.5 KiB | 23.54 MiB | 76.23 MiB | 678,667 |
| Create and save | Large | 218.84 ms | 90.07 MiB | 3.20 MiB | 70.59 MiB | 146.92 MiB | 3,350,485 |
| Load and inspect | Small | 4.07 ms | 1.45 MiB | 46.4 KiB | 46.42 MiB | 91.74 MiB | 47,522 |
| Load and inspect | Normal | 38.02 ms | 15.78 MiB | 728.9 KiB | 48.72 MiB | 109.34 MiB | 678,667 |
| Load and inspect | Large | 226.83 ms | 93.04 MiB | 3.56 MiB | 70.09 MiB | 141.20 MiB | 3,350,485 |

The isolated runner includes validation and process startup behavior outside its
timed loop, while BenchmarkDotNet supplies the primary steady-state time and
allocation measurements. Large creation still allocates roughly 27 times the
resulting package size, and large loading allocates roughly 29 times the input
size. Those ratios are owner-level optimization signals, not comparisons with
another implementation.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- --filter '*VisioBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\visio\evidence.json
```

Raw BenchmarkDotNet and process evidence remain ignored machine-local
artifacts. This note retains the compact reproducible result and exact source
provenance.
