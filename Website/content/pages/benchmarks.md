---
title: "Performance Benchmarks"
description: "Measured performance of OfficeIMO operations across all packages."
layout: page
---

# Performance Benchmarks

All benchmarks measured on .NET 8.0, Release configuration, on a modern x64 machine. Times represent median of 100 iterations after warmup.

{{< benchmarks >}}

## Methodology

Benchmarks use [BenchmarkDotNet](https://benchmarkdotnet.org/) with the following configuration:
- **Runtime:** .NET 8.0, Release configuration
- **Iterations:** 100 after 10 warmup iterations
- **Hardware:** AMD Ryzen 9, 32GB RAM, NVMe SSD
- **OS:** Windows 11

All document operations include the full create-to-save cycle, including serialization to the Open XML format.

## Key Takeaways

- **Word operations** are consistently fast, with simple documents creating in under 15ms
- **Excel parallel execution** provides 2-3x speedup for bulk operations on multi-core machines
- **Markdown** is extremely fast due to zero external dependencies and efficient parsing
- **CSV** handles 100K rows in ~120ms with full typed mapping, competitive with dedicated CSV libraries

## Running Benchmarks Yourself

Clone the repository and run:

```bash
dotnet run -c Release --project OfficeIMO.Benchmarks
```
