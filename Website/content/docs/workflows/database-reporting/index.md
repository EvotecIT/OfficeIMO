---
title: "Database Reporting with DbaClientX"
description: "Move SQL data into Excel or CSV, verify the artifact, and optionally load it back with DbaClientX and PSWriteOffice."
order: 6
meta.seo_title: "Export SQL Server data to Excel or CSV with PowerShell"
---

DbaClientX owns database access and bulk writes. OfficeIMO owns the document and data formats. PSWriteOffice connects the two in a PowerShell pipeline without duplicating either engine.

## SQL query to an Excel report

```powershell
Import-Module DbaClientX
Import-Module PSWriteOffice

$rows = Invoke-DbaXQuery `
    -Server 'sql01' `
    -Database 'Operations' `
    -Query @'
SELECT Department, Amount, CreatedUtc
FROM dbo.MonthlyRevenue
ORDER BY Department;
'@ `
    -ReturnType PSObject

$rows | Export-OfficeExcel `
    -Path '.\Monthly-Revenue.xlsx' `
    -WorksheetName 'Revenue' `
    -TableName 'MonthlyRevenue' `
    -BoldTopRow `
    -FreezeTopRow `
    -AutoFit
```

This route is useful for scheduled operational reports because the query layer stays replaceable and the workbook layer remains independently testable.

## Verify the workbook and load it back

```powershell
$table = Import-OfficeExcel `
    -Path '.\Monthly-Revenue.xlsx' `
    -WorksheetName 'Revenue' `
    -AsDataTable

if ($table.Rows.Count -eq 0) {
    throw 'The generated workbook contains no report rows.'
}

$table | Write-DbaXTableData `
    -Provider SqlServer `
    -ConnectionString $connectionString `
    -DestinationTable 'dbo.MonthlyRevenueImport' `
    -AutoCreateTable `
    -BatchSize 5000
```

The repository includes a complete [Excel round-trip example](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Excel/Example-ExcelDbaClientXRoundTrip.ps1) that creates source data, exports a workbook, imports a `DataTable`, writes it to a destination table, verifies row counts, and removes its test artifacts.

## CSV for large or interoperable handoffs

Replace the Excel commands with `Export-OfficeCsv` and `Import-OfficeCsv` when a delimited file is the better contract:

```powershell
$rows | Export-OfficeCsv -Path '.\Monthly-Revenue.csv'
$table = Import-OfficeCsv -Path '.\Monthly-Revenue.csv' -AsDataTable -InferSchema
```

The [CSV round-trip example](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Csv/Example-CsvDbaClientXRoundTrip.ps1) verifies the same database-to-file-to-database flow. For a streaming database handoff, use `Import-OfficeCsv -AsDataReader` and pipe the reader to `Write-DbaXTableData`; dispose reader and database resources when the transfer completes.

## Production checklist

1. Parameterize the query and keep credentials in the database connection layer.
2. Select columns explicitly so workbook and CSV schemas remain stable.
3. Validate row counts and required columns before publishing or loading data back.
4. Use `DataTable` when the next step needs typed, buffered tabular data; use `IDataReader` for large streaming transfers.
5. Store the generated artifact, validation result, and query/run identity together when reports are audit evidence.

See [PSWriteOffice performance evidence](/docs/workflows/powershell-benchmarks/) for the Excel and CSV benchmark methodology used by these paths.
