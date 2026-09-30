<#
.SYNOPSIS
Creates Microsoft Excel pivot ranking fixtures with zero and negative measures.
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
        [pscustomobject]@{ Key = 'zero-top-percent'; Type = 3; Value = 40.0; Amounts = @(0,10,20,30,40,50); Total = 90.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'zero-bottom-percent'; Type = 4; Value = 40.0; Amounts = @(0,10,20,30,40,50); Total = 60.0; Range = 'A4:B9' },
        [pscustomobject]@{ Key = 'zero-top-sum'; Type = 5; Value = 60.0; Amounts = @(0,10,20,30,40,50); Total = 90.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'zero-bottom-sum'; Type = 6; Value = 60.0; Amounts = @(0,10,20,30,40,50); Total = 60.0; Range = 'A4:B9' },
        [pscustomobject]@{ Key = 'mixed-top-count'; Type = 1; Value = 1.0; Amounts = @(-40,-10,0,20,30,50); Total = 50.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'mixed-bottom-count'; Type = 2; Value = 1.0; Amounts = @(-40,-10,0,20,30,50); Total = -40.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'mixed-top-percent'; Type = 3; Value = 40.0; Amounts = @(-40,-10,0,20,30,50); Total = 50.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'mixed-bottom-percent'; Type = 4; Value = 40.0; Amounts = @(-40,-10,0,20,30,50); Total = 50.0; Range = 'A4:B11' },
        [pscustomobject]@{ Key = 'mixed-top-sum'; Type = 5; Value = 60.0; Amounts = @(-40,-10,0,20,30,50); Total = 80.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'mixed-bottom-sum'; Type = 6; Value = 20.0; Amounts = @(-40,-10,0,20,30,50); Total = 50.0; Range = 'A4:B11' },
        [pscustomobject]@{ Key = 'negative-top-percent'; Type = 3; Value = 40.0; Amounts = @(-60,-50,-40,-30,-20,-10); Total = -100.0; Range = 'A4:B9' },
        [pscustomobject]@{ Key = 'negative-bottom-percent'; Type = 4; Value = 40.0; Amounts = @(-60,-50,-40,-30,-20,-10); Total = -110.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'mixed-negative-top-percent'; Type = 3; Value = 40.0; Amounts = @(30,-10,-15,-20,-20,-25); Total = -35.0; Range = 'A4:B10' },
        [pscustomobject]@{ Key = 'mixed-negative-bottom-percent'; Type = 4; Value = 40.0; Amounts = @(30,-10,-15,-20,-20,-25); Total = -25.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'balanced-top-percent'; Type = 3; Value = 40.0; Amounts = @(-50,-20,0,10,20,40); Total = 0.0; Range = 'A4:B11' },
        [pscustomobject]@{ Key = 'balanced-bottom-percent'; Type = 4; Value = 40.0; Amounts = @(-50,-20,0,10,20,40); Total = 0.0; Range = 'A4:B11' }
    )
    $names = @('Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot')
    foreach ($case in $cases) {
        $workbook = $application.Workbooks.Add()
        $source = $workbook.Worksheets.Item(1)
        $source.Name = 'Source'
        $source.Cells.Item(1, 1).Value2 = 'Region'
        $source.Cells.Item(1, 2).Value2 = 'Sales'
        for ($index = 0; $index -lt $names.Count; $index++) {
            $source.Cells.Item($index + 2, 1).Value2 = $names[$index]
            $source.Cells.Item($index + 2, 2).Value2 = [double]$case.Amounts[$index]
        }
        $view = $workbook.Worksheets.Add()
        $view.Name = 'Grouped'
        $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R7C2", 6)
        $pivot = $cache.CreatePivotTable($view.Range('A4'), 'ValuePivot')
        $field = $pivot.PivotFields('Region')
        $field.Orientation = 1
        $field.Position = 1
        $measure = $pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
        $pivot.RowAxisLayout(1)
        [void]$pivot.RefreshTable()
        [void]$field.PivotFilters.Add2($case.Type, $measure, $case.Value)
        $lookups = $workbook.Worksheets.Add()
        $lookups.Name = 'Lookups'
        $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
        for ($index = 0; $index -lt $names.Count; $index++) {
            $lookups.Cells.Item($index + 2, 2).Formula =
                '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","' + $names[$index] + '")'
        }
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = $pivot.GetPivotData('Metric').Value2
        if ($grand -ne $case.Total -or $range -ne $case.Range) {
            throw "Unexpected Excel $($case.Key) result: range=$range grand=$grand"
        }
        $file = "pivot-value-$($case.Key)-conformance.xlsx"
        $path = Join-Path $directory $file
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
        $workbook = $null
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotNonpositiveRankingOracle.ps1'
            file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = 'Source!A1:B7'; amounts = $case.Amounts
            filterType = $case.Type; value1 = $case.Value
            outputRange = $range; grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-value-$($case.Key)-conformance.provenance.json") -Encoding utf8
        [pscustomobject]@{ File = $file; Range = $range; Grand = $grand }
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
