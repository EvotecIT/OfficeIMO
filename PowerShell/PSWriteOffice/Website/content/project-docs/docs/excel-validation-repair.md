---
title: "Validate and Repair Excel Workbooks"
description: "Preflight workbook operations, test integrity and accessibility, diagnose package state, and apply deliberate repairs."
layout: docs
---

Validation is a delivery stage, not an incidental side effect of opening a workbook. Keep inspection separate from repair so an automation run cannot silently rewrite a source file.

## Preflight first

- `Get-OfficeExcelRuntimePreflight` reports runtime capabilities and warnings.
- `Get-OfficeExcelPreflight` reports workbook-specific readiness and repair hints.
- `Test-OfficeExcelWorkbook` checks workbook integrity.
- `Test-OfficeExcelAccessibility` checks accessible-delivery concerns.
- `Get-OfficeExcelSummary -IncludeSchema` inventories package features such as queries, slicers, timelines, and external structures.

## Repair explicitly

Use `Repair-OfficeExcelWorkbook` only after inspecting the proposed repair class and preserving the original when required. Reopen the result, repeat the integrity checks, and compare it with the source if preservation matters.

For automated delivery, a useful gate is: create or update, save, preflight, test integrity, test accessibility when applicable, compare against a reference contract, then publish. The [operational dashboard](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1) provides a substantial artifact for exercising that path.
