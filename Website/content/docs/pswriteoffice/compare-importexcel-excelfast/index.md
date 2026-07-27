---
title: "PSWriteOffice vs ImportExcel vs ExcelFast"
description: "Compare PowerShell Excel workflows, supported scope, maturity, and reproducible performance evidence for PSWriteOffice, ImportExcel, and ExcelFast."
layout: docs
---

PSWriteOffice, ImportExcel, and ExcelFast can read or write XLSX files from PowerShell without automating Microsoft Excel. They are not interchangeable: each project optimizes for a different workflow and exposes a different public command surface.

This comparison uses the projects' public documentation and the reproducible benchmark suite in the PSWriteOffice repository. Upstream facts were last checked on 27 July 2026.

## Short answer

- Choose **ImportExcel** when its established `Import-Excel` and `Export-Excel` pipeline, examples, and existing script ecosystem already fit the job.
- Evaluate **ExcelFast** when its performance-focused `Import-Workbook`, `Export-Workbook`, `Get-Workbook`, and `Save-Workbook` model fits an experimental or early-stage workflow. Its own README currently describes the project as alpha.
- Choose **PSWriteOffice** when the automation needs Excel plus Word, PowerPoint, PDF, CSV, email, Visio, or open formats; or when workbook inspection, preflight, repair, accessibility, comparison, and repeatable report composition matter.

## Capability and project shape

| Question | PSWriteOffice | ImportExcel | ExcelFast |
| --- | --- | --- | --- |
| Primary interface | `OfficeExcel` document DSL and targeted commands | Object pipeline centered on `Import-Excel` and `Export-Excel` | Workbook import/export/edit commands |
| Microsoft Excel required | No | No | No |
| Platforms stated by project | PowerShell on supported .NET targets | Windows, Linux, and macOS | PowerShell module; confirm the target environment during evaluation |
| Tables and charts | Yes | Yes | Check the current alpha command surface for the required feature |
| Pivots | Yes | Yes | Check the current alpha command surface for the required feature |
| Workbook preflight, repair, comparison, and accessibility commands | Yes | Different surface; assess the exact required operation | Different surface; assess the exact required operation |
| Formats beyond Excel and CSV in the same module | Word, PowerPoint, PDF, email, Visio, Markdown, RTF, OpenDocument, and others | No; Excel-focused | No; Excel and CSV-focused |
| License | MIT | Apache-2.0 | MIT |

The table describes public project scope, not a promise that similarly named features preserve the same workbook structures. Test formulas, cached values, charts, pivots, named ranges, data validation, formatting, external links, macros, and malformed inputs that matter to the workload.

## Equivalent starting points

### Export PowerShell objects

ImportExcel keeps the common case compact:

```powershell
$rows | Export-Excel -Path '.\Report.xlsx' -WorksheetName 'Data' -AutoSize
```

PSWriteOffice makes the document structure explicit:

```powershell
New-OfficeExcel -Path '.\Report.xlsx' {
    Add-OfficeExcelSheet -Name 'Data' {
        Add-OfficeExcelTable -InputObject $rows -TableName 'Data' -AutoFit
    }
}
```

ExcelFast documents this entry point:

```powershell
$rows | Export-Workbook '.\Report.xlsx'
```

### Import worksheet rows

```powershell
$importExcelRows = Import-Excel -Path '.\Report.xlsx' -WorksheetName 'Data'
$psWriteOfficeRows = Import-OfficeExcel -Path '.\Report.xlsx' -WorksheetName 'Data'
$excelFastRows = Import-Workbook '.\Report.xlsx'
```

These examples are syntactically similar but may differ in type inference, formula handling, range behavior, and metadata retained. Compare the returned objects with the workbook shapes used in production.

## Reproducible performance evidence

PSWriteOffice includes a PowerForge benchmark matrix that runs comparable engines next to each other, alternates execution order, validates generated workbooks, and records skipped lanes instead of treating unsupported operations as wins. A committed benchmark snapshot includes:

| Scenario | Rows | PSWriteOffice | ExcelFast | ImportExcel |
| --- | ---: | ---: | ---: | ---: |
| Full default worksheet import | 10,000 | 189.6 ms | 226.6 ms | 520.5 ms |
| Text-object workbook write | 10,000 | 160.8 ms | 790.5 ms | 3.17 s |
| Update an existing workbook | 10,000 | 2.39 s | Not compared | 2.65 s |

These numbers are evidence for the recorded environment and exact scenarios, not a universal ranking. ExcelFast write lanes are limited to equivalent text-only shapes in the current suite, while mixed typed writes are excluded because the compared output is not equivalent. Repeat the benchmark on the deployment hardware and preserve workbook validation:

```powershell
$env:OfficeIMORoot = (Resolve-Path '..\OfficeIMO').Path

pwsh -NoProfile -File .\Benchmarks\Compare-ExcelPerformance.ps1 `
    -Suite Standard `
    -RowCount 1000,5000,10000 `
    -Engine PSWriteOffice,ImportExcel,ExcelFast
```

See the [benchmark methodology and complete result matrix](https://github.com/EvotecIT/PSWriteOffice/blob/main/Benchmarks/README.md) before quoting a result.

## Choose ImportExcel when

- existing automation already depends on its command and parameter model;
- the concise object-to-worksheet pipeline is the main requirement;
- its public examples cover the report shape and operational constraints;
- avoiding migration risk matters more than consolidating document formats.

See the official [ImportExcel repository](https://github.com/dfinke/ImportExcel) and [ImportExcel PowerShell Gallery package](https://www.powershellgallery.com/packages/ImportExcel).

## Choose ExcelFast when

- its streaming or workbook-editing model matches the workload;
- an alpha dependency is acceptable and tested in the deployment environment;
- the required workbook features are present in the current command surface;
- performance is evaluated with output validation rather than elapsed time alone.

See the official [ExcelFast repository](https://github.com/JustinGrote/ExcelFast) and [ExcelFast PowerShell Gallery package](https://www.powershellgallery.com/packages/ExcelFast).

## Choose PSWriteOffice when

- one script produces or inspects several document formats;
- the workbook workflow needs targeted reads and updates, templates, formulas, tables, charts, pivots, validation, comments, accessibility, comparison, or repair;
- PowerShell is the automation surface, while OfficeIMO remains available as the underlying .NET API;
- the team wants a checked-in benchmark it can rerun against its own data shapes.

Start with [Excel automation](/docs/pswriteoffice/excel/) or [choose a workflow](/docs/pswriteoffice/choosing-a-workflow/).
