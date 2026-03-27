---
title: "Performance Benchmarks"
description: "Representative performance snapshots for common OfficeIMO workflows."
layout: page
---

# Performance Benchmarks

The figures below are representative performance snapshots for common OfficeIMO scenarios. Treat them as directional guidance rather than formal cross-machine guarantees.

{{< benchmarks >}}

## How to Read These Numbers

- The scenarios are intended to show relative shape and typical workflow cost, not certify a universal SLA.
- Real throughput depends on document size, formatting complexity, runtime version, storage, CPU, and whether your workload is write-only or read/modify.
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

For example, the repo currently includes a dedicated benchmark project for Markdown:

```bash
dotnet run -c Release --project OfficeIMO.Markdown.Benchmarks
```
