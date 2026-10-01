<#
.SYNOPSIS
Creates Microsoft Excel pivot ranking fixtures with one error-valued measure.
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
        [pscustomobject]@{ Key = 'top-count'; Type = 1; Value = 1.0; Function = -4157; Total = 50.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'bottom-count'; Type = 2; Value = 1.0; Function = -4157; Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'top-percent'; Type = 3; Value = 40.0; Function = -4157; Total = 90.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'bottom-percent'; Type = 4; Value = 40.0; Function = -4157; Total = 60.0; Range = 'A4:B8' },
        [pscustomobject]@{ Key = 'top-sum'; Type = 5; Value = 60.0; Function = -4157; Total = 90.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'bottom-sum'; Type = 6; Value = 60.0; Function = -4157; Total = 60.0; Range = 'A4:B8' },
        [pscustomobject]@{ Key = 'average-top-count'; Type = 1; Value = 1.0; Function = -4106; Total = 50.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'count-top-count'; Type = 1; Value = 1.0; Function = -4112; Total = 6.0; Range = 'A4:B11' },
        [pscustomobject]@{ Key = 'count-numbers-bottom-count'; Type = 2; Value = 1.0; Function = -4113; Total = 0.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'div-zero-top-count'; Type = 1; Value = 1.0; Function = -4157; ErrorFormula = '=1/0'; Total = 50.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'value-bottom-sum'; Type = 6; Value = 60.0; Function = -4157; ErrorFormula = '=VALUE("bad")'; Total = 60.0; Range = 'A4:B8' },
        [pscustomobject]@{ Key = 'mixed-group-bottom-percent'; Type = 4; Value = 40.0; Function = -4157; ExtraBravo = $true; Total = 60.0; Range = 'A4:B8' }
    )
    $names = @('Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot')
    $amounts = @(10, $null, 20, 30, 40, 50)
    foreach ($case in $cases) {
        $errorFormula = if ($case.ErrorFormula) { $case.ErrorFormula } else { '=NA()' }
        $sourceEndRow = if ($case.ExtraBravo) { 8 } else { 7 }
        $workbook = $application.Workbooks.Add()
        $source = $workbook.Worksheets.Item(1)
        $source.Name = 'Source'
        $source.Cells.Item(1, 1).Value2 = 'Region'
        $source.Cells.Item(1, 2).Value2 = 'Sales'
        for ($index = 0; $index -lt $names.Count; $index++) {
            $source.Cells.Item($index + 2, 1).Value2 = $names[$index]
            if ($null -eq $amounts[$index]) {
                $source.Cells.Item($index + 2, 2).Formula = $errorFormula
            } else {
                $source.Cells.Item($index + 2, 2).Value2 = [double]$amounts[$index]
            }
        }
        if ($case.ExtraBravo) {
            $source.Cells.Item(8, 1).Value2 = 'Bravo'
            $source.Cells.Item(8, 2).Value2 = 15.0
        }
        $application.CalculateFullRebuild()
        $view = $workbook.Worksheets.Add()
        $view.Name = 'Grouped'
        $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R$($sourceEndRow)C2", 6)
        $pivot = $cache.CreatePivotTable($view.Range('A4'), 'ValuePivot')
        $field = $pivot.PivotFields('Region')
        $field.Orientation = 1
        $field.Position = 1
        $measure = $pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', $case.Function)
        $pivot.RowAxisLayout(1)
        [void]$pivot.RefreshTable()
        try {
            [void]$field.PivotFilters.Add2($case.Type, $measure, $case.Value)
        } catch {
            [pscustomobject]@{ Case = $case.Key; Error = $_.Exception.Message }
            $workbook.Close($false)
            $workbook = $null
            continue
        }
        $lookups = $workbook.Worksheets.Add()
        $lookups.Name = 'Lookups'
        $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
        for ($index = 0; $index -lt $names.Count; $index++) {
            $lookups.Cells.Item($index + 2, 2).Formula =
                '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","' + $names[$index] + '")'
        }
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = [double]$pivot.GetPivotData('Metric').Value2
        if ($grand -ne $case.Total -or $range -ne $case.Range) {
            throw "Unexpected Excel $($case.Key) result: range=$range grand=$grand"
        }
        $displayedItems = @()
        for ($viewRow = 5; $viewRow -le 10; $viewRow++) {
            $label = [string]$view.Cells.Item($viewRow, 1).Text
            if ($label -and $label -ne 'Grand Total') { $displayedItems += $label }
        }
        $file = "pivot-value-error-$($case.Key)-conformance.xlsx"
        $path = Join-Path $directory $file
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
        $workbook = $null
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotErrorRankingOracle.ps1'
            file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = "Source!A1:B$sourceEndRow"; errorCells = 'Source!B3'; errorFormula = $errorFormula
            filterType = $case.Type; value1 = $case.Value; aggregateFunction = $case.Function; outputRange = $range
            grandTotal = $grand
            displayedItems = $displayedItems
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-value-error-$($case.Key)-conformance.provenance.json") -Encoding utf8
        [pscustomobject]@{ Case = $case.Key; Range = $range; Grand = $grand; Displayed = ($displayedItems -join ',') }
    }
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $application) {
        $application.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
