---
title: "Validate and Repair Excel Workbooks"
description: "Preflight workbook operations, test integrity and accessibility, diagnose package state, and apply deliberate repairs."
layout: docs
---

Validation is a delivery stage, not an incidental side effect of opening a workbook. Keep inspection separate from repair so an automation run cannot silently rewrite a source file.

## Preflight first

- `Import-OfficeExcel -Path .\Report.xlsx -Culture en-US -WarningAction Stop` surfaces runtime import warnings as errors.
- `Test-OfficeDocument -Path .\Report.xlsx` checks workbook integrity.
- Add `-Options (New-OfficeExcelTestOptions -Capability EditCellValues,ExportPdfReport -IncludeFeatures -IncludeRepairHints)` to evaluate the intended operations and retain their repair hints.
- `Test-OfficeDocument -Path .\Report.xlsx -Check Accessibility` checks accessible-delivery concerns.
- `Get-OfficeDocumentInfo -Path .\Report.xlsx -Detailed` returns common counts and a native workbook snapshot in `Details`.

Validation returns `IsValid`, `Diagnostics`, and `NativeReports`. Inspect the native reports for workbook doctor and capability-preflight findings; use `-Quiet` only when the consumer needs a Boolean.

## Repair explicitly

Use `Repair-OfficeDocument` only after inspecting the proposed repair class and preserving the original when required. Reopen the result, repeat the integrity checks, and compare it with the source if preservation matters.

For automated delivery, a useful gate is: create or update, save, preflight, test integrity, test accessibility when applicable, compare against a reference contract, then publish. The [operational dashboard](https://github.com/EvotecIT/OfficeIMO/blob/945cb285f954b745c0f4ebfce0849c0726a76fa1/PowerShell/PSWriteOffice/Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1) provides a substantial artifact for exercising that path.
