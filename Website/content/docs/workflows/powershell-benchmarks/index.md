---
title: "PSWriteOffice Performance Evidence"
description: "Reproduce the Excel and CSV benchmarks used to evaluate PSWriteOffice PowerShell workflows."
order: 8
meta.seo_title: "PSWriteOffice Excel and CSV benchmarks"
---

PSWriteOffice uses the PSPublishModule/PowerForge benchmark runner so scenario rotation, cleanup, result validation, and report generation share one implementation. Excel is compared with PowerShell-facing workbook tools; CSV is compared with native PowerShell CSV handling.

## What is compared

- Every Excel reader receives the same PSWriteOffice-generated workbook shape.
- Unsupported or semantically different competitor lanes are marked `Skipped`, not treated as wins.
- CSV comparisons include object, `DataTable`, compression, wide-column, quoted-field, and dbatools-shaped workloads.
- Output validation happens inside the scenario contract. Managed cleanup runs after setup and outside the timed operation.
- Results are local snapshots with retained run context, not promises for every machine or workload.

## Run the smoke suites

From a PSWriteOffice checkout with a sibling OfficeIMO source tree:

```powershell
$env:OfficeIMORoot = (Resolve-Path ..\OfficeIMO).Path

pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass `
    -File .\Benchmarks\Compare-ExcelPerformance.ps1 `
    -Suite Smoke

pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass `
    -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Smoke
```

Use `-Plan` to inspect the selected scenarios without running them. Use `-UpdateReadme` only when intentionally recording a new result snapshot.

## Selected committed results

These rows are examples from the current committed PSWriteOffice benchmark report. They deliberately include a result where native PowerShell is faster.

| Workload | Rows | PSWriteOffice | Comparison | Result |
|---|---:|---:|---:|---|
| Excel `report-workbook` | 10,000 | 889.0 ms | ImportExcel 2.09 s | PSWriteOffice fastest |
| Excel `import-default-full` | 10,000 | 189.6 ms | ExcelFast 226.6 ms; ImportExcel 520.5 ms | PSWriteOffice fastest |
| CSV `DataTable` write | 10,000 | 59.8 ms | NativeCsv 42.2 ms | NativeCsv fastest |
| CSV `DataTable` read, mixed | 100,000 | 219.6 ms | NativeCsv 2.57 s | PSWriteOffice fastest |
| CSV GZip write, wide | 10,000 | 118.8 ms | NativeCsv 397.7 ms | PSWriteOffice fastest |

Do not choose a package from one row. Run the scenario matching your row count, schema, compression, formatting, and validation needs on infrastructure close to production.

Open the [complete benchmark source and generated tables](https://github.com/EvotecIT/PSWriteOffice/tree/main/Benchmarks) or return to the [OfficeIMO benchmark hub](/benchmarks/).
