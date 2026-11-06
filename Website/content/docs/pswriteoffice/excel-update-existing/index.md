---
title: "Update Existing Excel Workbooks"
description: "Replace values, edit rows, append table data, and make targeted workbook changes without rebuilding every sheet."
layout: docs
---

Existing-workbook automation should express the smallest useful mutation. That makes the script easier to review and reduces the chance of replacing formulas, styles, or unrelated worksheets.

## Replace text or edit rows

`Update-OfficeExcelText` supports sheet and range boundaries, literal or regex matching, case sensitivity, and `-WhatIf`. `Edit-OfficeExcelRow` exposes a row object with header-aware access for conditional changes.

```powershell
Update-OfficeExcelText -Path '.\Readiness.xlsx' -Sheet Readiness `
    -Pattern Draft -Replacement Ready -SimpleMatch

Edit-OfficeExcelRow -Path '.\Readiness.xlsx' -Sheet Readiness -ScriptBlock {
    param($row)
    if ($row.CellByHeader('Service').Value -eq 'Messaging') {
        $row.Set('Owner', 'Productivity')
    }
}
```

## Append without rebuilding

Open the workbook with `Open-OfficeExcel`, pipe it to `Add-OfficeExcelTableRow`, then save with `Save-OfficeExcel` and close with `Close-OfficeExcel`. Other targeted commands update cells, formulas, styles, links, print settings, visibility, active sheet, filters, comments, validation, and workbook metadata.

The [update-existing recipe](https://github.com/EvotecIT/OfficeIMO/blob/3b3a029106ece280577eca40744bd196d6c7b5c4/PowerShell/PSWriteOffice/Examples/Excel/Recipe-Excel-UpdateExisting.ps1) combines workbook-wide text replacement with a header-aware row edit.
