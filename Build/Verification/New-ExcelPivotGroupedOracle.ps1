<#
.SYNOPSIS
Creates an independently calculated Excel pivot with numeric range grouping.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
$path = Join-Path $directory 'numeric-group-conformance.xlsx'
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
        @('Quantity', 'Sales'),
        @(-2.0, 1.0), @(4.0, 40.0), @(7.0, 70.0), @(12.0, 120.0),
        @(18.0, 180.0), @(22.0, 220.0), @(29.0, 290.0), @(35.0, 2.0)
    )
    $values = New-Object 'object[,]' $rows.Count, 2
    for ($row = 0; $row -lt $rows.Count; $row++) {
        for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
    }
    $range = $source.Range('A1:B9')
    try { $range.Value2 = $values } finally { Release-Com $range }
    $view = $workbook.Worksheets.Add()
    $view.Name = 'Grouped'
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R9C2", 6)
    $destination = $view.Range('A4')
    try { $pivot = $cache.CreatePivotTable($destination, 'PivotGrouped') }
    finally { Release-Com $destination }
    $quantity = $pivot.PivotFields('Quantity')
    $quantity.Orientation = 1
    $quantity.Position = 1
    $sales = $pivot.PivotFields('Sales')
    $measure = $pivot.AddDataField($sales, 'Metric', -4157)
    $pivot.RowAxisLayout(1)
    [void]$pivot.RefreshTable()
    $groupCell = $quantity.DataRange.Cells.Item(1, 1)
    try { [void]$groupCell.Group(0, 30, 10) }
    finally { Release-Com $groupCell }
    $queries = @(
        @('Grand', '=GETPIVOTDATA("Metric",Grouped!A4)'),
        @('Below', '=GETPIVOTDATA("Metric",Grouped!A4,"Quantity","<0")'),
        @('First', '=GETPIVOTDATA("Metric",Grouped!A4,"Quantity","0-9")'),
        @('Second', '=GETPIVOTDATA("Metric",Grouped!A4,"Quantity","10-19")'),
        @('Third', '=GETPIVOTDATA("Metric",Grouped!A4,"Quantity","20-30")'),
        @('Above', '=GETPIVOTDATA("Metric",Grouped!A4,"Quantity",">30")'),
        @('Raw', '=GETPIVOTDATA("Metric",Grouped!A4,"Quantity",4)')
    )
    $lookups = $workbook.Worksheets.Add()
    $lookups.Name = 'Lookups'
    for ($row = 0; $row -lt $queries.Count; $row++) {
        $label = $lookups.Cells.Item($row + 1, 1)
        $formula = $lookups.Cells.Item($row + 1, 2)
        try { $label.Value2 = $queries[$row][0]; $formula.Formula = $queries[$row][1] }
        finally { Release-Com $formula; Release-Com $label }
    }
    $application.CalculateFullRebuild()
    $tableRange = $pivot.TableRange1
    try { $outputRange = $tableRange.Address($false, $false) }
    finally { Release-Com $tableRange }
    $actual = @()
    for ($row = 1; $row -le $queries.Count; $row++) {
        $cell = $lookups.Cells.Item($row, 2)
        try { $actual += [ordered]@{ name = $queries[$row - 1][0]; value = $cell.Text } }
        finally { Release-Com $cell }
    }
    $workbook.SaveAs($path, 51)
    $firstChanged = $source.Range('A2')
    $secondChanged = $source.Range('A3')
    try { $firstChanged.Value2 = 15.0; $secondChanged.Value2 = 31.0 }
    finally { Release-Com $secondChanged; Release-Com $firstChanged }
    [void]$pivot.RefreshTable()
    $application.CalculateFullRebuild()
    $refreshedPath = Join-Path $directory 'numeric-group-refresh-conformance.xlsx'
    $refreshedRange = $pivot.TableRange1
    try { $refreshedOutputRange = $refreshedRange.Address($false, $false) }
    finally { Release-Com $refreshedRange }
    $refreshedLookups = @()
    for ($row = 1; $row -le $queries.Count; $row++) {
        $cell = $lookups.Cells.Item($row, 2)
        try { $refreshedLookups += [ordered]@{ name = $queries[$row - 1][0]; value = $cell.Text } }
        finally { Release-Com $cell }
    }
    $workbook.SaveAs($refreshedPath, 51)
    $workbook.Close($false)
    $closed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotGroupedOracle.ps1'
        file = 'numeric-group-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B9'; outputRange = $outputRange
        grouping = 'Quantity:0-30 by 10'; lookups = $actual
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'numeric-group-conformance.provenance.json') -Encoding utf8
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotGroupedOracle.ps1'
        file = 'numeric-group-refresh-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $refreshedPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B9'; outputRange = $refreshedOutputRange
        grouping = 'Quantity:0-30 by 10'; changedSource = 'A2=-2 to 15; A3=4 to 31'
        lookups = $refreshedLookups
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'numeric-group-refresh-conformance.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($value in @($lookups, $measure, $sales, $quantity, $pivot, $cache, $view, $source, $workbook, $application)) {
        Release-Com $value
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
