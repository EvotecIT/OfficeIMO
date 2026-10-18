---
title: "Merge and Compare Excel Workbooks"
description: "Consolidate selected worksheets and produce explicit workbook or range differences for review and delivery gates."
layout: docs
---

Workbook consolidation and workbook comparison solve different problems. Joining copies selected sheets into a target workbook. Comparing reports how two workbooks differ without combining them.

## Consolidate selected sheets

```powershell
Join-OfficeExcelWorkbook -Path '.\Consolidated.xlsx' `
    -SourcePath '.\Regions.xlsx' `
    -SourceSheet North,South `
    -SheetNamePrefix 'Region '
```

Use a stable prefix when sources can contain the same sheet names. Copy mode and name-validation options let the operation match the preservation and naming policy.

## Compare a candidate

`Compare-OfficeExcelWorkbook` checks cells, styles, named ranges, tables, worksheet metadata, and comments. Skip categories only when the delivery contract says they do not matter. `Compare-OfficeExcelRange` narrows the comparison to a known business area.

The [merge-workbooks recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-MergeWorkbooks.ps1) consolidates two regional sheets. The [compare-workbooks recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Recipe-Excel-CompareWorkbooks.ps1) creates a controlled cell difference and reports it.

After a merge, run the checks described in [validation and repair](/docs/pswriteoffice/excel-validation-repair/).
