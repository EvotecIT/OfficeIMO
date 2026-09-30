<#
.SYNOPSIS
Creates an Excel-calculated pivot with a report filter and a hidden row item.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
$path = Join-Path $directory 'filtered-conformance.xlsx'
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
        @('Region', 'Product', 'Amount'),
        @('North', 'A', 10.0), @('North', 'B', 20.0),
        @('South', 'A', 5.0), @('South', 'B', 15.0),
        @('West', 'B', 7.0), @('North', 'B', 3.0)
    )
    $values = New-Object 'object[,]' $rows.Count, 3
    for ($row = 0; $row -lt $rows.Count; $row++) {
        for ($column = 0; $column -lt 3; $column++) { $values[$row, $column] = $rows[$row][$column] }
    }
    $range = $source.Range('A1:C7')
    try { $range.Value2 = $values } finally { Release-Com $range }
    $view = $workbook.Worksheets.Add()
    $view.Name = 'Filtered'
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R7C3", 6)
    $destination = $view.Range('A4')
    try { $pivot = $cache.CreatePivotTable($destination, 'PivotFiltered') }
    finally { Release-Com $destination }
    $region = $pivot.PivotFields('Region')
    $region.Orientation = 1
    $region.Position = 1
    $region.Subtotals = @($false, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false)
    $product = $pivot.PivotFields('Product')
    $product.Orientation = 3
    $product.Position = 1
    $amount = $pivot.PivotFields('Amount')
    $measure = $pivot.AddDataField($amount, 'Metric', -4157)
    $pivot.RowAxisLayout(1)
    $pivot.RowGrand = $true
    $pivot.ColumnGrand = $true
    [void]$pivot.RefreshTable()
    $product.CurrentPage = 'B'
    $west = $region.PivotItems('West')
    $west.Visible = $false
    $queries = @(
        @('Grand', '=GETPIVOTDATA("Metric",Filtered!A4)'),
        @('North', '=GETPIVOTDATA("Metric",Filtered!A4,"Region","North")'),
        @('South', '=GETPIVOTDATA("Metric",Filtered!A4,"Region","South")'),
        @('West', '=GETPIVOTDATA("Metric",Filtered!A4,"Region","West")'),
        @('ProductB', '=GETPIVOTDATA("Metric",Filtered!A4,"Product","B")'),
        @('ProductA', '=GETPIVOTDATA("Metric",Filtered!A4,"Product","A")'),
        @('NorthB', '=GETPIVOTDATA("Metric",Filtered!A4,"Region","North","Product","B")')
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

    # Excel refresh is the oracle for whether a newly discovered row item joins
    # a manually filtered field and whether a page selection survives reindexing.
    $newRegion = $source.Range('A2')
    $newProduct = $source.Range('B2')
    try { $newRegion.Value2 = 'East'; $newProduct.Value2 = 'B' }
    finally { Release-Com $newProduct; Release-Com $newRegion }
    $eastLookup = $lookups.Range('A8:B8')
    try { $eastLookup.Cells.Item(1, 1).Value2 = 'East'; $eastLookup.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Filtered!A4,"Region","East")' }
    finally { Release-Com $eastLookup }
    [void]$pivot.RefreshTable()
    $application.CalculateFullRebuild()
    $refreshedPath = Join-Path $directory 'filtered-refresh-conformance.xlsx'
    $refreshedRange = $pivot.TableRange1
    try { $refreshedOutputRange = $refreshedRange.Address($false, $false) }
    finally { Release-Com $refreshedRange }
    $refreshedLookups = @()
    for ($row = 1; $row -le 8; $row++) {
        $cell = $lookups.Cells.Item($row, 2)
        try { $refreshedLookups += [ordered]@{ name = $(if ($row -eq 8) { 'East' } else { $queries[$row - 1][0] }); value = $cell.Text } }
        finally { Release-Com $cell }
    }
    $workbook.SaveAs($refreshedPath, 51)

    $blankProduct = $source.Range('B3')
    try { $blankProduct.Value2 = $null }
    finally { Release-Com $blankProduct }
    [void]$pivot.RefreshTable()
    $product.CurrentPage = '(blank)'
    $blankLookup = $lookups.Range('A9:B9')
    try { $blankLookup.Cells.Item(1, 1).Value2 = 'ProductBlank'; $blankLookup.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Filtered!A4,"Product","(blank)")' }
    finally { Release-Com $blankLookup }
    $application.CalculateFullRebuild()
    $blankPath = Join-Path $directory 'filtered-blank-conformance.xlsx'
    $blankRange = $pivot.TableRange1
    try { $blankOutputRange = $blankRange.Address($false, $false) }
    finally { Release-Com $blankRange }
    $blankLookups = @()
    for ($row = 1; $row -le 9; $row++) {
        $cell = $lookups.Cells.Item($row, 2)
        try { $blankLookups += [ordered]@{ name = $(if ($row -eq 9) { 'ProductBlank' } elseif ($row -eq 8) { 'East' } else { $queries[$row - 1][0] }); value = $cell.Text } }
        finally { Release-Com $cell }
    }
    $workbook.SaveAs($blankPath, 51)
    $workbook.Close($false)
    $closed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotFilteredOracle.ps1'
        file = 'filtered-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:C7'; outputRange = $outputRange
        selectedPageItem = 'Product=B'; hiddenRowItem = 'Region=West'; lookups = $actual
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'filtered-conformance.provenance.json') -Encoding utf8

    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotFilteredOracle.ps1'
        file = 'filtered-refresh-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $refreshedPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:C7'; outputRange = $refreshedOutputRange
        selectedPageItem = 'Product=B'; hiddenRowItem = 'Region=West'; newRowItem = 'Region=East'
        lookups = $refreshedLookups
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'filtered-refresh-conformance.provenance.json') -Encoding utf8

    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotFilteredOracle.ps1'
        file = 'filtered-blank-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $blankPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:C7'; outputRange = $blankOutputRange
        selectedPageItem = 'Product=(blank)'; hiddenRowItem = 'Region=West'; newRowItem = 'Region=East'
        lookups = $blankLookups
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'filtered-blank-conformance.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($value in @($lookups, $measure, $amount, $west, $product, $region, $pivot, $cache, $view, $source, $workbook, $application)) {
        Release-Com $value
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
