<#
.SYNOPSIS
Creates an Excel-calculated pivot with manually grouped text items.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
$path = Join-Path $directory 'manual-group-conformance.xlsx'
$refreshPath = Join-Path $directory 'manual-group-refresh-conformance.xlsx'
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
        @('Product', 'Sales'),
        @('Apple', 10.0), @('Pear', 20.0), @('Carrot', 30.0),
        @('Broccoli', 40.0), @('Apple', 5.0)
    )
    $values = New-Object 'object[,]' $rows.Count, 2
    for ($row = 0; $row -lt $rows.Count; $row++) {
        for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
    }
    $sourceRange = $source.Range('A1:B6')
    try { $sourceRange.Value2 = $values } finally { Release-Com $sourceRange }
    $view = $workbook.Worksheets.Add()
    $view.Name = 'Grouped'
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R6C2", 6)
    $destination = $view.Range('A4')
    try { $pivot = $cache.CreatePivotTable($destination, 'PivotManualGrouped') }
    finally { Release-Com $destination }
    $product = $pivot.PivotFields('Product')
    $product.Orientation = 1
    $product.Position = 1
    $sales = $pivot.PivotFields('Sales')
    $measure = $pivot.AddDataField($sales, 'Metric', -4157)
    $pivot.RowAxisLayout(1)
    [void]$pivot.RefreshTable()
    $apple = $product.PivotItems('Apple').LabelRange
    $pear = $product.PivotItems('Pear').LabelRange
    $selection = $application.Union($apple, $pear)
    try { [void]$selection.Group() }
    finally { Release-Com $selection; Release-Com $pear; Release-Com $apple }
    $parent = $pivot.PivotFields('Product2')
    $parent.PivotItems('Group1').Name = 'Fruit'
    $lookups = $workbook.Worksheets.Add()
    $lookups.Name = 'Lookups'
    $queries = @(
        @('Grand', '=GETPIVOTDATA("Metric",Grouped!$A$4)'),
        @('Fruit', '=GETPIVOTDATA("Metric",Grouped!$A$4,"Product2","Fruit")'),
        @('Apple', '=GETPIVOTDATA("Metric",Grouped!$A$4,"Product2","Fruit","Product","Apple")'),
        @('Pear', '=GETPIVOTDATA("Metric",Grouped!$A$4,"Product2","Fruit","Product","Pear")'),
        @('Carrot', '=GETPIVOTDATA("Metric",Grouped!$A$4,"Product2","Carrot")'),
        @('Unknown', '=GETPIVOTDATA("Metric",Grouped!$A$4,"Product2","Unknown")')
    )
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
    $workbook.SaveAs($path, 51)
    $changed = $source.Range('A2')
    try { $changed.Value2 = 'Carrot' } finally { Release-Com $changed }
    [void]$pivot.RefreshTable()
    $application.CalculateFullRebuild()
    $refreshedRange = $pivot.TableRange1
    try { $refreshOutputRange = $refreshedRange.Address($false, $false) }
    finally { Release-Com $refreshedRange }
    $workbook.SaveAs($refreshPath, 51)
    $workbook.Close($false)
    $closed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotManualGroupOracle.ps1'
        file = 'manual-group-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B6'; outputRange = $outputRange
        grouping = 'Apple and Pear manually grouped as Fruit'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'manual-group-conformance.provenance.json') -Encoding utf8
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotManualGroupOracle.ps1'
        file = 'manual-group-refresh-conformance.xlsx'; sha256 = (Get-FileHash -LiteralPath $refreshPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B6'; outputRange = $refreshOutputRange
        grouping = 'Apple and Pear manually grouped as Fruit'
        mutation = 'A2 changed from Apple to Carrot'
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'manual-group-refresh-conformance.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($value in @($lookups, $parent, $measure, $sales, $product, $pivot, $cache, $view, $source, $workbook, $application)) {
        Release-Com $value
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
