---
title: "Create an Excel workbook"
description: "Use PSWriteOffice to create an Excel workbook with a summary table."
layout: docs
---

This pattern is useful when a PowerShell script should produce a workbook instead of only console output.

It is adapted from `Examples/Excel/Example-ExcelBasic.ps1`.

## Example

```powershell
Import-Module PSWriteOffice

$outputPath = Join-Path $PSScriptRoot 'Output\RevenueSnapshot.xlsx'
$null = New-Item -ItemType Directory -Path (Split-Path $outputPath) -Force
$data = @(
    [PSCustomObject]@{ Region = 'North America'; Revenue = 125000; YoY = 0.12 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000; YoY = 0.22 }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000; YoY = 0.18 }
)

New-OfficeExcel -Path $outputPath {
    Add-OfficeExcelSheet -Name 'Summary' -Content {
        Set-OfficeExcelRow -Row 1 -Values 'Region', 'Revenue', 'YoY'
        Add-OfficeExcelTable -Data $data -TableName 'Sales' -TableStyle 'TableStyleMedium9'
        Set-OfficeExcelColumn -Column 1 -AutoFit
    }
}
```

## What this demonstrates

- creating an Excel workbook without automating desktop Excel
- writing a named worksheet and table
- keeping generated files under a local output folder

## Source

- [Example-ExcelBasic.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Example-ExcelBasic.ps1)
