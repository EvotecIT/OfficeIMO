<#
.SYNOPSIS
Creates an Excel-calculated pivot with years and months date grouping.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
$path = Join-Path $directory 'date-group-conformance.xlsx'
$refreshPath = Join-Path $directory 'date-group-refresh-conformance.xlsx'
$valuesFirstPath = Join-Path $directory 'date-group-values-first-conformance.xlsx'
$valuesMiddlePath = Join-Path $directory 'date-group-values-middle-conformance.xlsx'
$valuesLastPath = Join-Path $directory 'date-group-values-last-conformance.xlsx'
$columnsPath = Join-Path $directory 'date-group-columns-conformance.xlsx'
$mutex = [Threading.Mutex]::new($false, 'Local\OfficeIMO.Excel.Tests.DesktopCom')
$acquired = $false
$application = $null
$workbook = $null
$closed = $false
function Release-Com($value) {
    if ($null -ne $value) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) }
}
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
        @([DateTime]::new(2026, 1, 10).ToOADate(), 40.0),
        @([DateTime]::new(2026, 4, 5).ToOADate(), 50.0)
    )
    $values = New-Object 'object[,]' $rows.Count, 2
    for ($row = 0; $row -lt $rows.Count; $row++) {
        for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
    }
    $range = $source.Range('A1:B6')
    try { $range.Value2 = $values } finally { Release-Com $range }
    $dateRange = $source.Range('A2:A6')
    try { $dateRange.NumberFormat = 'yyyy-mm-dd' } finally { Release-Com $dateRange }
    $view = $workbook.Worksheets.Add()
    $view.Name = 'Grouped'
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R6C2", 6)
    $destination = $view.Range('A4')
    try { $pivot = $cache.CreatePivotTable($destination, 'PivotDateGrouped') }
    finally { Release-Com $destination }
    $dateField = $pivot.PivotFields('OrderDate')
    $dateField.Orientation = 1
    $dateField.Position = 1
    $sales = $pivot.PivotFields('Sales')
    $measure = $pivot.AddDataField($sales, 'Metric', -4157)
    $pivot.RowAxisLayout(1)
    [void]$pivot.RefreshTable()
    $groupCell = $dateField.DataRange.Cells.Item(1, 1)
    $periods = @($false, $false, $false, $false, $true, $false, $true)
    try { [void]$groupCell.Group($true, $true, 1, $periods) }
    finally { Release-Com $groupCell }
    $lookups = $workbook.Worksheets.Add()
    $lookups.Name = 'Lookups'
    $formulas = @(
        '=GETPIVOTDATA("Metric",Grouped!$A$4)',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)",2025)',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)","2025")',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)",2025,"Months (OrderDate)","sty")',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)","2026","Months (OrderDate)","sty")',
        '=GETPIVOTDATA("Metric",Grouped!$A$4,"Years (OrderDate)","2025","Months (OrderDate)","mar")'
    )
    for ($index = 0; $index -lt $formulas.Count; $index++) {
        $formulaCell = $lookups.Cells.Item($index + 1, 2)
        try { $formulaCell.Formula = $formulas[$index] }
        finally { Release-Com $formulaCell }
    }
    $application.CalculateFullRebuild()
    $tableRange = $pivot.TableRange1
    try { $outputRange = $tableRange.Address($false, $false) }
    finally { Release-Com $tableRange }
    $workbook.SaveAs($path, 51)
    $changed = $source.Range('A3:A4')
    try {
        $changedValues = New-Object 'object[,]' 2, 1
        $changedValues[0, 0] = [DateTime]::new(2025, 1, 15).ToOADate()
        $changedValues[1, 0] = [DateTime]::new(2025, 1, 15).ToOADate()
        $changed.Value2 = $changedValues
    } finally { Release-Com $changed }
    [void]$pivot.RefreshTable()
    $application.CalculateFullRebuild()
    $tableRange = $pivot.TableRange1
    try { $refreshOutputRange = $tableRange.Address($false, $false) }
    finally { Release-Com $tableRange }
    $workbook.SaveAs($refreshPath, 51)
    $workbook.Close($false)
    $closed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotDateGroupOracle.ps1'
        file = 'date-group-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B6'; outputRange = $outputRange; grouping = 'Years and Months'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'date-group-conformance.provenance.json') -Encoding utf8
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotDateGroupOracle.ps1'
        file = 'date-group-refresh-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $refreshPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B6'; outputRange = $refreshOutputRange; grouping = 'Years and Months'
        mutation = 'OrderDate A3 and A4 changed to 2025-01-15, contracting five date keys to three'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'date-group-refresh-conformance.provenance.json') -Encoding utf8

    $valuesWorkbook = $application.Workbooks.Open($path)
    $valuesView = $valuesWorkbook.Worksheets.Item('Grouped')
    $valuesPivot = $valuesView.PivotTables('PivotDateGrouped')
    $valuesSales = $valuesPivot.PivotFields('Sales')
    $countMeasure = $valuesPivot.AddDataField($valuesSales, 'Count', -4112)
    $valuesAxis = $valuesPivot.DataPivotField
    $valuesAxis.Orientation = 1
    $valuesAxis.Position = 1
    $valuesPivot.RowAxisLayout(1)
    [void]$valuesPivot.RefreshTable()
    $application.CalculateFullRebuild()
    $valuesRange = $valuesPivot.TableRange1
    try { $valuesOutputRange = $valuesRange.Address($false, $false) }
    finally { Release-Com $valuesRange }
    $valuesWorkbook.SaveAs($valuesFirstPath, 51)
    $valuesWorkbook.Close($false)
    $valuesClosed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotDateGroupOracle.ps1'
        file = 'date-group-values-first-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $valuesFirstPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B6'; outputRange = $valuesOutputRange; grouping = 'Years and Months'
        layout = 'Values first on rows; Sum and Count measures'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'date-group-values-first-conformance.provenance.json') -Encoding utf8

    $middleWorkbook = $application.Workbooks.Open($valuesFirstPath)
    $middleView = $middleWorkbook.Worksheets.Item('Grouped')
    $middlePivot = $middleView.PivotTables('PivotDateGrouped')
    $middleAxis = $middlePivot.DataPivotField
    $middleAxis.Position = 2
    $middlePivot.RowAxisLayout(1)
    [void]$middlePivot.RefreshTable()
    $application.CalculateFullRebuild()
    $middleRange = $middlePivot.TableRange1
    try { $middleOutputRange = $middleRange.Address($false, $false) }
    finally { Release-Com $middleRange }
    $middleWorkbook.SaveAs($valuesMiddlePath, 51)
    $middleWorkbook.Close($false)
    $middleClosed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotDateGroupOracle.ps1'
        file = 'date-group-values-middle-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $valuesMiddlePath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B6'; outputRange = $middleOutputRange; grouping = 'Years and Months'
        layout = 'Values between Years and Months on rows; Sum and Count measures'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'date-group-values-middle-conformance.provenance.json') -Encoding utf8

    $lastWorkbook = $application.Workbooks.Open($valuesMiddlePath)
    $lastView = $lastWorkbook.Worksheets.Item('Grouped')
    $lastPivot = $lastView.PivotTables('PivotDateGrouped')
    $lastAxis = $lastPivot.DataPivotField
    $lastAxis.Position = 3
    $lastPivot.RowAxisLayout(1)
    [void]$lastPivot.RefreshTable()
    $application.CalculateFullRebuild()
    $lastRange = $lastPivot.TableRange1
    try { $lastOutputRange = $lastRange.Address($false, $false) }
    finally { Release-Com $lastRange }
    $lastWorkbook.SaveAs($valuesLastPath, 51)
    $lastWorkbook.Close($false)
    $lastClosed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotDateGroupOracle.ps1'
        file = 'date-group-values-last-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $valuesLastPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B6'; outputRange = $lastOutputRange; grouping = 'Years and Months'
        layout = 'Values after Years and Months on rows; Sum and Count measures'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'date-group-values-last-conformance.provenance.json') -Encoding utf8

    $columnWorkbook = $application.Workbooks.Open($path)
    $columnView = $columnWorkbook.Worksheets.Item('Grouped')
    $columnPivot = $columnView.PivotTables('PivotDateGrouped')
    $yearField = $columnPivot.PivotFields('Years (OrderDate)')
    $monthField = $columnPivot.PivotFields('Months (OrderDate)')
    $yearField.Orientation = 2
    $yearField.Position = 1
    $monthField.Orientation = 2
    $monthField.Position = 2
    [void]$columnPivot.RefreshTable()
    $application.CalculateFullRebuild()
    $columnRange = $columnPivot.TableRange1
    try { $columnOutputRange = $columnRange.Address($false, $false) }
    finally { Release-Com $columnRange }
    $columnWorkbook.SaveAs($columnsPath, 51)
    $columnWorkbook.Close($false)
    $columnClosed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotDateGroupOracle.ps1'
        file = 'date-group-columns-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $columnsPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B6'; outputRange = $columnOutputRange; grouping = 'Years and Months'
        layout = 'Years and Months on columns; Sum measure'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'date-group-columns-conformance.provenance.json') -Encoding utf8

} finally {
    if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
    if ($null -ne $valuesWorkbook -and -not $valuesClosed) { $valuesWorkbook.Close($false) }
    if ($null -ne $middleWorkbook -and -not $middleClosed) { $middleWorkbook.Close($false) }
    if ($null -ne $lastWorkbook -and -not $lastClosed) { $lastWorkbook.Close($false) }
    if ($null -ne $columnWorkbook -and -not $columnClosed) { $columnWorkbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($value in @($monthField, $yearField, $columnPivot, $columnView, $columnWorkbook, $lastAxis, $lastPivot, $lastView, $lastWorkbook, $middleAxis, $middlePivot, $middleView, $middleWorkbook, $countMeasure, $valuesSales, $valuesAxis, $valuesPivot, $valuesView, $valuesWorkbook, $measure, $sales, $dateField, $pivot, $cache, $lookups, $view, $source, $workbook, $application)) {
        Release-Com $value
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
