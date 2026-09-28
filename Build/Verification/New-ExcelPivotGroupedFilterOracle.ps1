<#
.SYNOPSIS
Creates independently calculated Excel pivot-item filter fixtures for qualified groups.
#>
[CmdletBinding()]
param([switch] $RenamedOnly)
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
$mutex = [Threading.Mutex]::new($false, 'Local\OfficeIMO.Excel.Tests.DesktopCom')
$acquired = $false
$application = $null
try {
    $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5))
    if (-not $acquired) { throw 'Excel COM lock timed out.' }
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    if (-not $RenamedOnly) {
    $scenarios = @(
        [pscustomobject]@{ Source = 'manual-group-conformance.xlsx'; Output = 'manual-group-filter-conformance.xlsx'; Pivot = 'PivotManualGrouped'; Hides = @(@('Product2', 'Carrot'), @('Product', 'Apple')) },
        [pscustomobject]@{ Source = 'numeric-group-conformance.xlsx'; Output = 'numeric-group-filter-conformance.xlsx'; Pivot = 'PivotGrouped'; Hides = @(,@('Quantity', '0-9')) },
        [pscustomobject]@{ Source = 'date-group-conformance.xlsx'; Output = 'date-group-filter-conformance.xlsx'; Pivot = 'PivotDateGrouped'; Hides = @(,@('Years (OrderDate)', '2026')) }
    )
    foreach ($scenario in $scenarios) {
        $workbook = $null
        try {
            $sourcePath = Join-Path $directory $scenario.Source
            $outputPath = Join-Path $directory $scenario.Output
            $workbook = $application.Workbooks.Open($sourcePath, 0, $false)
            $sheet = $workbook.Worksheets.Item('Grouped')
            $pivot = $sheet.PivotTables($scenario.Pivot)
            foreach ($hide in $scenario.Hides) {
                $pivot.PivotFields($hide[0]).PivotItems($hide[1]).Visible = $false
            }
            $application.CalculateFullRebuild()
            $range = $pivot.TableRange1.Address($false, $false)
            $grand = $pivot.GetPivotData('Metric').Value2
            $workbook.SaveAs($outputPath, 51)
            $workbook.Close($false)
            $workbook = $null
            [ordered]@{
                producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
                generatedUtc = [DateTime]::UtcNow.ToString('o')
                regeneration = 'Build/Verification/New-ExcelPivotGroupedFilterOracle.ps1'
                source = $scenario.Source; file = $scenario.Output
                sha256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash.ToLowerInvariant()
                hiddenItems = @($scenario.Hides | ForEach-Object { "$($_[0]):$($_[1])" })
                outputRange = $range; grandTotal = $grand
            } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory ($scenario.Output -replace '\.xlsx$', '.provenance.json')) -Encoding utf8
            [pscustomobject]@{ File = $scenario.Output; Range = $range; Grand = $grand }
        } finally {
            if ($null -ne $workbook) { $workbook.Close($false) }
        }
    }
    $workbook = $null
    try {
        $sourcePath = Join-Path $directory 'manual-group-filter-conformance.xlsx'
        $outputPath = Join-Path $directory 'manual-group-filter-refresh-conformance.xlsx'
        $workbook = $application.Workbooks.Open($sourcePath, 0, $false)
        $source = $workbook.Worksheets.Item('Source')
        $source.Range('A2').Value2 = 'Broccoli'
        $pivot = $workbook.Worksheets.Item('Grouped').PivotTables('PivotManualGrouped')
        [void]$pivot.RefreshTable()
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = $pivot.GetPivotData('Metric').Value2
        $workbook.SaveAs($outputPath, 51)
        $workbook.Close($false)
        $workbook = $null
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotGroupedFilterOracle.ps1'
            source = 'manual-group-filter-conformance.xlsx'; file = 'manual-group-filter-refresh-conformance.xlsx'
            sha256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash.ToLowerInvariant()
            mutation = 'Source!A2 changed from hidden Apple to visible Broccoli, then refreshed'
            outputRange = $range; grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'manual-group-filter-refresh-conformance.provenance.json') -Encoding utf8
        [pscustomobject]@{ File = 'manual-group-filter-refresh-conformance.xlsx'; Range = $range; Grand = $grand }
    } finally {
        if ($null -ne $workbook) { $workbook.Close($false) }
    }
    }
    $renames = @(
        [pscustomobject]@{ Source = 'numeric-group-filter-conformance.xlsx'; Output = 'numeric-group-filter-renamed-conformance.xlsx'; Pivot = 'PivotGrouped'; Field = 'Quantity'; Item = '10-19'; Caption = 'Small' },
        [pscustomobject]@{ Source = 'date-group-filter-conformance.xlsx'; Output = 'date-group-filter-renamed-conformance.xlsx'; Pivot = 'PivotDateGrouped'; Field = 'Years (OrderDate)'; Item = '2025'; Caption = 'FY25' }
    )
    foreach ($scenario in $renames) {
        $workbook = $null
        try {
            $sourcePath = Join-Path $directory $scenario.Source
            $outputPath = Join-Path $directory $scenario.Output
            $workbook = $application.Workbooks.Open($sourcePath, 0, $false)
            $pivot = $workbook.Worksheets.Item('Grouped').PivotTables($scenario.Pivot)
            $pivot.PivotFields($scenario.Field).PivotItems($scenario.Item).Caption = $scenario.Caption
            $application.CalculateFullRebuild()
            $range = $pivot.TableRange1.Address($false, $false)
            $grand = $pivot.GetPivotData('Metric').Value2
            $renamedLookup = $pivot.GetPivotData('Metric', $scenario.Field, $scenario.Caption).Value2
            $workbook.SaveAs($outputPath, 51)
            $workbook.Close($false)
            $workbook = $null
            [ordered]@{
                producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
                generatedUtc = [DateTime]::UtcNow.ToString('o')
                regeneration = 'Build/Verification/New-ExcelPivotGroupedFilterOracle.ps1'
                source = $scenario.Source; file = $scenario.Output
                sha256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash.ToLowerInvariant()
                renamedItem = "$($scenario.Field):$($scenario.Item) -> $($scenario.Caption)"
                outputRange = $range; grandTotal = $grand; renamedLookup = $renamedLookup
            } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory ($scenario.Output -replace '\.xlsx$', '.provenance.json')) -Encoding utf8
            [pscustomobject]@{ File = $scenario.Output; Range = $range; Grand = $grand }
        } finally {
            if ($null -ne $workbook) { $workbook.Close($false) }
        }
    }
} finally {
    if ($null -ne $application) { $application.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
