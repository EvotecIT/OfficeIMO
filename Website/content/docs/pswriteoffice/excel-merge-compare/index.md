---
title: "Merge and Compare Excel Workbooks"
description: "Consolidate selected worksheets and produce explicit workbook or range differences for review and delivery gates."
layout: docs
---

Copy selected sheets into an existing target workbook, then compare the result with a baseline.

## Consolidate selected sheets

Start with an existing target workbook, or create one with `New-OfficeExcel`.

```powershell
Copy-OfficeExcelSheet -Path '.\Consolidated.xlsx' `
    -Source '.\Regions.xlsx' `
    -Sheet North,South `
    -NamePrefix 'Region '
```

Use a stable prefix when sources can contain the same sheet names. Copy mode and name-validation options let the operation match the preservation and naming policy.

## Compare a candidate

`Compare-OfficeDocument -ReferencePath .\Before.xlsx -DifferencePath .\After.xlsx` checks cells, styles, named ranges, tables, worksheet metadata, and comments. The common report exposes `AreEqual`, `Differences`, and `NativeReport`. Pass `New-OfficeExcelCompareOptions` through `-Options` to choose the comparison scope. Skip categories only when the delivery contract says they do not matter. `Compare-OfficeExcelRange` narrows the comparison to a known business area.

The [merge-workbooks recipe](https://github.com/EvotecIT/OfficeIMO/blob/8562059074dcc0fa4df5db990f93469a7e4a00d6/PowerShell/PSWriteOffice/Examples/Excel/Recipe-Excel-MergeWorkbooks.ps1) consolidates two regional sheets. The [compare-workbooks recipe](https://github.com/EvotecIT/OfficeIMO/blob/8562059074dcc0fa4df5db990f93469a7e4a00d6/PowerShell/PSWriteOffice/Examples/Excel/Recipe-Excel-CompareWorkbooks.ps1) creates a controlled cell difference and reports it.

After a merge, run the checks described in [validation and repair](/docs/pswriteoffice/excel-validation-repair/).
