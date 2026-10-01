<#
.SYNOPSIS
Creates a Microsoft Excel workbook with two views of one native pivot cache.
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
    $values = New-Object 'object[,]' 5, 2
    $rows = @(@('Region', 'Sales'), @('East', 10.0), @('West', 20.0), @('East', 30.0), @('North', 40.0))
    for ($row = 0; $row -lt 5; $row++) {
        for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
    }
    $source.Range('A1:B5').Value2 = $values
    $rowsView = $workbook.Worksheets.Add()
    $rowsView.Name = 'Rows'
    $columnsView = $workbook.Worksheets.Add()
    $columnsView.Name = 'Columns'
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R5C2", 6)
    $first = $cache.CreatePivotTable($rowsView.Range('A4'), 'RowsPivot')
    $first.PivotFields('Region').Orientation = 1
    [void]$first.AddDataField($first.PivotFields('Sales'), 'Metric', -4157)
    $first.RowAxisLayout(1)
    [void]$first.RefreshTable()
    $second = $cache.CreatePivotTable($columnsView.Range('B5'), 'ColumnsPivot')
    $second.PivotFields('Region').Orientation = 2
    [void]$second.AddDataField($second.PivotFields('Sales'), 'Metric', -4157)
    [void]$second.RefreshTable()
    $application.CalculateFullRebuild()
    $rowTotal = $first.GetPivotData('Metric').Value2
    $columnTotal = $second.GetPivotData('Metric').Value2
    if ($rowTotal -ne 100.0 -or $columnTotal -ne 100.0) { throw "Unexpected pivot totals: $rowTotal, $columnTotal" }
    $cacheIndex = $first.CacheIndex
    $rowRange = $first.TableRange1.Address($false, $false)
    $columnRange = $second.TableRange1.Address($false, $false)
    $file = 'pivot-shared-cache-views-conformance.xlsx'
    $path = Join-Path $directory $file
    $workbook.SaveAs($path, 51)
    $workbook.Close($false)
    $workbook = $null
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelSharedPivotCacheOracle.ps1'
        file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B5'; pivotNames = @('RowsPivot', 'ColumnsPivot')
        pivotCacheIndex = $cacheIndex; rowOutputRange = $rowRange
        columnOutputRange = $columnRange
        rowGrandTotal = $rowTotal; columnGrandTotal = $columnTotal
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'pivot-shared-cache-views-conformance.provenance.json') -Encoding utf8
    [pscustomobject]@{ File = $file; Rows = $rowTotal; Columns = $columnTotal; Cache = $cacheIndex }
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $application) {
        $application.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
