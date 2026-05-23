---
title: "OfficeIMO Performance Benchmarks"
description: "Repeatable benchmark results showing OfficeIMO.Excel against common .NET Excel libraries."
layout: page
---

{{< benchmarks >}}

## Reproduce

```bash
dotnet run -c Release --framework net8.0 --project OfficeIMO.Excel.Benchmarks -- --compare-libraries Docs/benchmarks/comparison-current
.\Build\Generate-ExcelBenchmarkWebsiteData.ps1 -SummaryPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-summary.json -ManifestPath .\Docs\benchmarks\comparison-current\officeimo.excel.comparison-suite-manifest.json -RunMode quick
```
