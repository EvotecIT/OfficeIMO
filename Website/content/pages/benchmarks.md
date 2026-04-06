---
title: "Performance Benchmarks"
description: "Representative performance snapshots for common OfficeIMO workflows."
layout: page
---

# Performance Benchmarks

The figures below are representative performance snapshots for common OfficeIMO scenarios. Treat them as directional guidance rather than formal cross-machine guarantees.

{{< benchmarks >}}

## Excel Snapshot Status

The committed Excel snapshots in this repo are intentionally honest:

- `OfficeIMO.Excel` writes the sampled 2,500-row report scenario slightly faster than `ClosedXML` on the measured machine in the current committed snapshot.
- The current write-stage profile shows `InsertObjects()` as the largest staged cost, with `AutoFitColumns()` reduced to a smaller share of the write path.
- The committed Excel JSON artifacts include raw samples and medians, so outliers remain visible instead of being flattened into a single average.

## How to Read These Numbers

- The scenarios are intended to show relative shape and typical workflow cost, not certify a universal SLA.
- Real throughput depends on document size, formatting complexity, runtime version, storage, CPU, and whether your workload is write-only or read/modify.
- The Excel snapshot artifacts record multiple samples and medians. If the average and median diverge sharply, treat the run as noisy and re-measure on your target hardware.
- If performance matters for your use case, benchmark your own document patterns rather than relying on generic sample numbers.

## What the Table Is Good For

- Spotting which packages are lightweight for report-generation style workloads.
- Understanding that Markdown and CSV are cheaper to process than full Open XML document pipelines.
- Seeing where parallel Excel operations can help on multi-core machines.

## Reproducing or Extending Benchmarks

Benchmark projects in this repo are package-specific rather than one universal harness. If you want to validate a workload:

1. Start from the relevant package or sample project.
2. Recreate the document shapes that matter to your application.
3. Measure on the target runtime, OS, and hardware you plan to ship.

For example, the repo currently includes dedicated benchmark projects for Markdown and Excel:

```bash
dotnet run -c Release --project OfficeIMO.Markdown.Benchmarks
dotnet run -c Release --framework net8.0 --project OfficeIMO.Excel.Benchmarks
```

And the Excel harness can emit committed artifacts for both end-to-end snapshots and write-path profiling:

```bash
dotnet run -c Release --framework net8.0 --project OfficeIMO.Excel.Benchmarks -- --snapshot Docs/benchmarks/officeimo.excel.snapshot.json
dotnet run -c Release --framework net8.0 --project OfficeIMO.Excel.Benchmarks -- --profile-write Docs/benchmarks/officeimo.excel.write-profile.json
```
