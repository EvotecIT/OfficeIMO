# Visio load, inspection, and package-buffer allocation evidence (2026-09-22)

## Result

The Visio loader shares the order tokens for modeled cells, and its preservation
entries no longer require a separate object for each child. Inspection snapshots
reuse immutable empty collections and project single Shape Data rows directly.
The create/save path now materializes bytes from the finalized package stream
without copying the complete VSDX through a second memory stream. These changes
reduce allocations for both benchmark workflows without changing package size.

## Load and inspect

The comparison uses clean source at baseline commit
`456fe132b9f015e5cce67500dd8987e5fb6f4b9a` and candidate commit
`50ab5b3180d539f53ee897b4d256dc1d8926ec1b`. Each revision ran three
isolated child-process measurements per workload on the same host. The table
reports the median allocated bytes per large load-and-inspect operation.

| Host | Baseline | Candidate | Reduction | VSDX bytes, both revisions |
| --- | ---: | ---: | ---: | ---: |
| Windows 11, Ryzen 9 9950X3D2, .NET SDK 10.0.112 | 65.03 MiB | 59.97 MiB | 7.8% | 3,350,485 |
| Ubuntu 24.04.3 WSL2, x64, .NET SDK 10.0.112 | 65.23 MiB | 59.97 MiB | 8.1% | 3,350,479 |
| macOS 27.0, Apple M4, .NET SDK 10.0.401 | 65.03 MiB | 59.97 MiB | 7.8% | 3,350,481 |

Package bytes differ slightly across hosts, but match between the two source
revisions on each host. The macOS runner reported zero for process working-set
peak; that metric is unavailable on this host and is not used here.

## Create and save

The package-buffer comparison uses clean source at baseline commit
`0f8249112db9dab17d5d8dda6e8c433d7aa881f0` and candidate commit
`26cd501ba3f58e50a00d73bb02c279da91c40a2a`. Each revision ran five isolated
large create/save child probes per host. The table reports medians.

| Host | Baseline allocation | Candidate allocation | Reduction | Baseline ms | Candidate ms | VSDX bytes, both revisions |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Windows 11, Ryzen 9 9950X3D2, .NET SDK 10.0.112 | 52.20 MiB | 49.00 MiB | 6.1% | 89.30 | 74.20 | 3,350,485 |
| Ubuntu 24.04.3 WSL2, x64, .NET SDK 10.0.112 | 52.20 MiB | 49.00 MiB | 6.1% | 70.81 | 73.79 | 3,350,479 |
| macOS 27.0, Apple M4, .NET SDK 10.0.112 | 52.20 MiB | 49.00 MiB | 6.1% | 138.86 | 131.20 | 3,350,481 |

Allocation fell by the same 3.19 MiB on every host. Windows and macOS elapsed
medians improved, while the Linux median increased by 4.2%, so this follow-up
makes no cross-platform speed claim. Stream saves remain staged before the
caller-owned destination is changed, and file saves still publish through the
atomic file helper after package generation succeeds.

## Validated contract

The deterministic corpus covers 1 page with 25 shapes, 4 pages with 400
shapes, and 8 pages with 2,000 shapes. The large package has 1,992 connectors
and 2,000 Shape Data rows. `validate` reopens every package and checks page,
shape, connector, Shape Data, and boundary-text content. Validation passed on
all three hosts for both source revisions. On Windows, the full Visio test
suite passed: 974 tests each on .NET 8 and .NET 10, and 972 on .NET Framework
4.7.2. An independent read-only review found no actionable issue in the
candidate's preservation or snapshot behavior.

## Reproduce

From a clean checkout of either commit:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- validate
dotnet run -c Release -f net10.0 --no-build --project .\OfficeIMO.Visio.Benchmarks -- evidence --repeat 3 --json .\.benchmark-artifacts\visio\evidence.json
```

The isolated runner records its source commit, dirty-tree state, runtime,
operating system, allocation, managed-heap growth, process peak where
available, and output bytes. Keep raw JSON and build output in a task-owned
ignored artifact directory and remove them after summarizing the run.

XML parsing and emission, preservation of unknown content, remaining model and
package allocations, and stable elapsed-time and peak-memory evidence remain
open optimization work.
