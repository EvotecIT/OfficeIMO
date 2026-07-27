---
title: "Migrate from PSWriteExcel to PSWriteOffice"
description: "Replace archived PSWriteExcel scripts with maintained PSWriteOffice commands for XLSX creation, import, formulas, charts, and validation."
layout: docs
---

[PSWriteExcel](https://github.com/EvotecIT/PSWriteExcel) is archived and replaced by PSWriteOffice. New scripts should use the `OfficeExcel` command family over OfficeIMO.Excel.

The current Excel DSL is more explicit than the legacy pipeline helpers. A simple `ConvertTo-Excel` export normally becomes `New-OfficeExcel` with a sheet and table block; scripts that update existing workbooks should use `Get-OfficeExcel`, targeted setters, and `Save-OfficeExcel`.

## Common command replacements

| PSWriteExcel | PSWriteOffice | Migration note |
| --- | --- | --- |
| `New-ExcelDocument` | `New-OfficeExcel` | Use `-NoSave` only when the document must remain open for later operations. |
| `Get-ExcelDocument` | `Get-OfficeExcel` | Inspect worksheets, ranges, tables, formulas, comments, and workbook metadata with targeted commands. |
| `Add-ExcelWorkSheet` | `Add-OfficeExcelSheet` | Add it inside a `New-OfficeExcel` or active workbook context. |
| `Add-ExcelWorkSheetCell` | `Set-OfficeExcelCell` | Address cells by A1 address or row and column. |
| `Add-ExcelWorksheetData` | `Add-OfficeExcelTable` or `Add-OfficeExcelDataSet` | Choose a table for object rows or the dataset command for `System.Data.DataSet`. |
| `ConvertTo-Excel` | `New-OfficeExcel` plus `Add-OfficeExcelTable` | The replacement is a document workflow, not a parameter-compatible pipeline alias. |
| `ConvertFrom-Excel` | `Import-OfficeExcel` | Select the sheet and range required by the consuming script. |
| `Set-ExcelWorkSheetAutoFit` | `Invoke-OfficeExcelAutoFit` | Apply it to the intended sheet or range. |
| `Set-ExcelWorkSheetFreezePane` | `Set-OfficeExcelFreeze` | Recheck pane coordinates after migration. |
| `Set-ExcelWorkSheetAutoFilter` | `Set-OfficeExcelAutoFilter` | Tables can also own their filtering behavior. |
| `Save-ExcelDocument` | `Save-OfficeExcel` | `New-OfficeExcel` saves automatically unless `-NoSave` is used. |

## Before: PSWriteExcel

```powershell
Get-Process |
    Select-Object Name, Id, CPU |
    ConvertTo-Excel -Path '.\Output\Processes.xlsx' -WorkSheetName 'Processes' -AutoFilter
```

## After: PSWriteOffice

```powershell
$processes = Get-Process |
    Select-Object Name, Id, CPU

New-OfficeExcel -Path '.\Output\Processes.xlsx' {
    Add-OfficeExcelSheet -Name 'Processes' {
        Add-OfficeExcelTable -InputObject $processes -TableName 'Processes' -AutoFit
    }
}
```

To read workbook data:

```powershell
$rows = Import-OfficeExcel -Path '.\Output\Processes.xlsx' -WorksheetName 'Processes'
$rows | Where-Object CPU -gt 10
```

## Validate before replacing the old job

- compare property-to-column mapping, null values, dates, number formats, and culture-sensitive values;
- recalculate or mark formulas for recalculation deliberately with `Save-OfficeExcel` options;
- verify filters, tables, styles, freeze panes, charts, and print settings;
- use `Test-OfficeExcelWorkbook` and, when relevant, `Test-OfficeExcelAccessibility`;
- benchmark representative large exports rather than assuming either object pipeline has the same memory profile.

See [Excel automation](/docs/pswriteoffice/excel/) and the generated [PowerShell command reference](/api/powershell/) for current syntax.
