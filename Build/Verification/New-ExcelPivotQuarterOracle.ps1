<#
.SYNOPSIS
Creates independent Excel Years/Quarters/Months pivot materialization fixtures.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
$mutex = [Threading.Mutex]::new($false, 'Local\OfficeIMO.Excel.Tests.DesktopCom')
$acquired = $false
$application = $null
$workbook = $null
$columnWorkbook = $null
try {
    $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5))
    if (-not $acquired) { throw 'Excel COM lock timed out.' }
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $workbook = $application.Workbooks.Add()
    $source = $workbook.Worksheets.Item(1)
    $source.Name = 'Source'
    $rows = @(
        @('OrderDate', 'Sales'),
        @([DateTime]::new(2025, 1, 15).ToOADate(), 10.0),
        @([DateTime]::new(2025, 3, 20).ToOADate(), 20.0),
        @([DateTime]::new(2025, 7, 1).ToOADate(), 30.0),
        @([DateTime]::new(2025, 11, 5).ToOADate(), 25.0),
        @([DateTime]::new(2026, 1, 10).ToOADate(), 40.0),
        @([DateTime]::new(2026, 4, 5).ToOADate(), 50.0),
        @([DateTime]::new(2026, 9, 9).ToOADate(), 60.0)
    )
    $values = New-Object 'object[,]' $rows.Count, 2
    for ($row = 0; $row -lt $rows.Count; $row++) {
        for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
    }
    $source.Range('A1:B8').Value2 = $values
    $source.Range('A2:A8').NumberFormat = 'yyyy-mm-dd'
    $view = $workbook.Worksheets.Add()
    $view.Name = 'Grouped'
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R8C2", 6)
    $pivot = $cache.CreatePivotTable($view.Range('A4'), 'PivotQuarterGrouped')
    $dateField = $pivot.PivotFields('OrderDate')
    $dateField.Orientation = 1
    $dateField.Position = 1
    [void]$pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
    $pivot.RowAxisLayout(1)
    [void]$pivot.RefreshTable()
    $periods = @($false, $false, $false, $false, $true, $true, $true)
    [void]$dateField.DataRange.Cells.Item(1, 1).Group($true, $true, 1, $periods)
    $lookups = $workbook.Worksheets.Add()
    $lookups.Name = 'Lookups'
    $formulas = @(
        '=GETPIVOTDATA("Metric",Grouped!$A$4)',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)",2025)',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)",2025,"Quarters (OrderDate)","Qtr1")',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)",2026,"Quarters (OrderDate)","Qtr3")',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)",2026,"Quarters (OrderDate)","Qtr4")'
    )
    for ($index = 0; $index -lt $formulas.Count; $index++) {
        $lookups.Cells.Item($index + 1, 2).Formula = $formulas[$index]
    }
    $application.CalculateFullRebuild()
    $rowPath = Join-Path $directory 'date-quarter-row-conformance.xlsx'
    $rowRange = $pivot.TableRange1.Address($false, $false)
    $rowGrand = $pivot.GetPivotData('Metric').Value2
    $workbook.SaveAs($rowPath, 51)
    $workbook.Close($false)
    $workbook = $null
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotQuarterOracle.ps1'
        file = 'date-quarter-row-conformance.xlsx'
        sha256 = (Get-FileHash -LiteralPath $rowPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B8'; grouping = 'Years, Quarters, Months'; layout = 'Row hierarchy'
        outputRange = $rowRange; grandTotal = $rowGrand
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'date-quarter-row-conformance.provenance.json') -Encoding utf8
    [pscustomobject]@{ File = 'date-quarter-row-conformance.xlsx'; Range = $rowRange; Grand = $rowGrand }

    $columnWorkbook = $application.Workbooks.Open($rowPath, 0, $false)
    $columnPivot = $columnWorkbook.Worksheets.Item('Grouped').PivotTables('PivotQuarterGrouped')
    foreach ($name in @('Years (OrderDate)', 'Quarters (OrderDate)', 'Months (OrderDate)')) {
        $field = $columnPivot.PivotFields($name)
        $field.Orientation = 2
        $field.Position = switch ($name) {
            'Years (OrderDate)' { 1 }
            'Quarters (OrderDate)' { 2 }
            default { 3 }
        }
    }
    [void]$columnPivot.RefreshTable()
    $application.CalculateFullRebuild()
    $columnPath = Join-Path $directory 'date-quarter-column-conformance.xlsx'
    $columnRange = $columnPivot.TableRange1.Address($false, $false)
    $columnGrand = $columnPivot.GetPivotData('Metric').Value2
    $columnWorkbook.SaveAs($columnPath, 51)
    $columnWorkbook.Close($false)
    $columnWorkbook = $null
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotQuarterOracle.ps1'
        file = 'date-quarter-column-conformance.xlsx'
        sha256 = (Get-FileHash -LiteralPath $columnPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B8'; grouping = 'Years, Quarters, Months'; layout = 'Column hierarchy'
        outputRange = $columnRange; grandTotal = $columnGrand
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'date-quarter-column-conformance.provenance.json') -Encoding utf8
    [pscustomobject]@{ File = 'date-quarter-column-conformance.xlsx'; Range = $columnRange; Grand = $columnGrand }
} finally {
    if ($null -ne $columnWorkbook) { $columnWorkbook.Close($false) }
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
