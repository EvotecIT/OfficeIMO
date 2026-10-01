<#
.SYNOPSIS
Creates independent Excel label, value, and combined pivot-filter fixtures.
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
    $cases = @(
        [pscustomobject]@{ Key = 'label'; Pivot = 'LabelPivot'; Label = $true; Value = $false; Rename = $false; Total = 115.0 },
        [pscustomobject]@{ Key = 'value'; Pivot = 'ValuePivot'; Label = $false; Value = $true; Rename = $false; Total = 170.0 },
        [pscustomobject]@{ Key = 'combined'; Pivot = 'CombinedPivot'; Label = $true; Value = $true; Rename = $false; Total = 100.0 },
        [pscustomobject]@{ Key = 'caption'; Pivot = 'CaptionPivot'; Label = $true; Value = $false; Rename = $true; Total = 145.0 }
    )
    foreach ($case in $cases) {
        $workbook = $application.Workbooks.Add()
        $source = $workbook.Worksheets.Item(1)
        $source.Name = 'Source'
        $rows = @(
            @('Region', 'Sales'), @('East', 25.0), @('East', 35.0),
            @('North', 10.0), @('North', 20.0), @('Northeast', 40.0),
            @('Southeast', 15.0), @('West', 70.0)
        )
        $values = New-Object 'object[,]' 8, 2
        for ($row = 0; $row -lt 8; $row++) {
            for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
        }
        $source.Range('A1:B8').Value2 = $values
        $view = $workbook.Worksheets.Add()
        $view.Name = 'Grouped'
        $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R8C2", 6)
        $pivot = $cache.CreatePivotTable($view.Range('A4'), $case.Pivot)
        $region = $pivot.PivotFields('Region')
        $region.Orientation = 1
        $region.Position = 1
        $metric = $pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
        $pivot.RowAxisLayout(1)
        [void]$pivot.RefreshTable()
        if ($case.Rename) { $region.PivotItems('North').Name = 'Easterly' }
        if ($case.Label) { [void]$region.PivotFilters.Add2(21, [Type]::Missing, 'east') }
        if ($case.Value) { [void]$region.PivotFilters.Add2(9, $metric, 30.0) }
        $lookups = $workbook.Worksheets.Add()
        $lookups.Name = 'Lookups'
        $names = @('North', 'Northeast', 'East', 'West', 'Southeast')
        $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
        for ($index = 0; $index -lt $names.Count; $index++) {
            $lookups.Cells.Item($index + 2, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","' + $names[$index] + '")'
        }
        $lookups.Cells.Item(7, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","Easterly")'
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = $pivot.GetPivotData('Metric').Value2
        if ($grand -ne $case.Total) { throw "Unexpected $($case.Key) grand total: $grand" }
        $file = "pivot-filter-$($case.Key)-conformance.xlsx"
        $path = Join-Path $directory $file
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
        $workbook = $null
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotFilterOracle.ps1'
            file = $file
            sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = 'Source!A1:B8'; filter = $case.Key
            outputRange = $range; grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-filter-$($case.Key)-conformance.provenance.json") -Encoding utf8
        [pscustomobject]@{ File = $file; Range = $range; Grand = $grand }
    }
    $combinedPath = Join-Path $directory 'pivot-filter-combined-conformance.xlsx'
    $workbook = $application.Workbooks.Open($combinedPath, 0, $false)
    $workbook.Worksheets.Item('Source').Range('B7').Value2 = 55.0
    $pivot = $workbook.Worksheets.Item('Grouped').PivotTables('CombinedPivot')
    [void]$pivot.RefreshTable()
    $application.CalculateFullRebuild()
    $range = $pivot.TableRange1.Address($false, $false)
    $grand = $pivot.GetPivotData('Metric').Value2
    if ($range -ne 'A4:B8' -or $grand -ne 155.0) {
        throw "Unexpected combined refresh result: $range, $grand"
    }
    $file = 'pivot-filter-combined-refresh-conformance.xlsx'
    $path = Join-Path $directory $file
    $workbook.SaveAs($path, 51)
    $workbook.Close($false)
    $workbook = $null
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotFilterOracle.ps1'
        file = $file
        sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B8'; filter = 'combined-refresh'
        sourceChange = 'Source!B7: 15 to 55'
        outputRange = $range; grandTotal = $grand
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'pivot-filter-combined-refresh-conformance.provenance.json') -Encoding utf8
    [pscustomobject]@{ File = $file; Range = $range; Grand = $grand }
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $application) {
        $application.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
