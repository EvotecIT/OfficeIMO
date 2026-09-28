<#
.SYNOPSIS
Creates Microsoft Excel pivot label-filter fixtures with formatted numeric keys.
#>
[CmdletBinding()]
param([string[]] $Kinds = @())
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
        [pscustomobject]@{ Key = 'contains-comma'; Type = 21; Criterion = ','; Format = '#,##0'; Total = 50.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'equals-grouped'; Type = 15; Criterion = '1,000'; Format = '#,##0'; Total = 20.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'not-equals-grouped'; Type = 16; Criterion = '1,000'; Format = '#,##0'; Total = 40.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'general-contains-one'; Type = 21; Criterion = '1'; Format = 'General'; Total = 30.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'midpoint-positive'; Type = 15; Criterion = '3'; Format = '#,##0'; Values = @('=5/2', '=9/2'); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'midpoint-negative'; Type = 15; Criterion = '-3'; Format = '#,##0'; Values = @('=-5/2', '=9/2'); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'decimal-two'; Type = 15; Criterion = '1.20'; Format = '0.00'; Values = @(1.2, 2.3); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'percent-one'; Type = 15; Criterion = '12.5%'; Format = '0.0%'; Values = @(0.125, 0.25); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'decimal-midpoint'; Type = 15; Criterion = '1.3'; Format = '0.0'; Values = @('=5/4', '=9/4'); Total = 10.0; Range = 'A4:B6' }
    )
    if (@($Kinds | Where-Object { $_ -notin $cases.Key }).Count -gt 0) {
        throw "Unknown pivot label fixture kind: $($Kinds -join ', ')"
    }
    foreach ($case in $cases) {
        if ($Kinds.Count -gt 0 -and $Kinds -notcontains $case.Key) { continue }
        $workbook = $application.Workbooks.Add()
        $source = $workbook.Worksheets.Item(1)
        $source.Name = 'Source'
        $source.Cells.Item(1, 1).Value2 = 'Item'
        $source.Cells.Item(1, 2).Value2 = 'Sales'
        $keys = if ($case.Values) { @($case.Values) } else { @(10, 1000, 2000) }
        $lastSourceRow = $keys.Count + 1
        for ($index = 0; $index -lt $keys.Count; $index++) {
            if ($keys[$index] -is [string] -and $keys[$index].StartsWith('=')) {
                $source.Cells.Item($index + 2, 1).Formula = $keys[$index]
            } else {
                $source.Cells.Item($index + 2, 1).Value2 = $keys[$index]
            }
            $source.Cells.Item($index + 2, 2).Value2 = [double](10 * ($index + 1))
        }
        if ($case.Format -ne 'General') { $source.Range("A2:A$lastSourceRow").NumberFormat = $case.Format }
        $application.CalculateFullRebuild()
        $view = $workbook.Worksheets.Add()
        $view.Name = 'Grouped'
        $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R$($lastSourceRow)C2", 6)
        $pivot = $cache.CreatePivotTable($view.Range('A4'), 'LabelPivot')
        $field = $pivot.PivotFields('Item')
        $field.Orientation = 1
        $field.Position = 1
        [void]$pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
        $pivot.RowAxisLayout(1)
        [void]$pivot.RefreshTable()
        [void]$field.PivotFilters.Add2($case.Type, [Type]::Missing, $case.Criterion)
        $lookups = $workbook.Worksheets.Add()
        $lookups.Name = 'Lookups'
        $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
        for ($index = 0; $index -lt $keys.Count; $index++) {
            $lookupKey = [string]$keys[$index]
            if ($lookupKey.StartsWith('=')) { $lookupKey = $lookupKey.Substring(1) }
            $lookups.Cells.Item($index + 2, 2).Formula =
                '=GETPIVOTDATA("Metric",Grouped!$A$4,"Item",' + $lookupKey + ')'
        }
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = [double]$pivot.GetPivotData('Metric').Value2
        if ($grand -ne $case.Total -or $range -ne $case.Range) {
            throw "Unexpected Excel $($case.Key) result: range=$range grand=$grand"
        }
        $file = "pivot-label-number-$($case.Key)-conformance.xlsx"
        $path = Join-Path $directory $file
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
        $workbook = $null
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotFormattedNumericLabelOracle.ps1'
            file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = "Source!A1:B$lastSourceRow"; numberFormat = $case.Format
            filterType = $case.Type; criterion = $case.Criterion
            outputRange = $range; grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-label-number-$($case.Key)-conformance.provenance.json") -Encoding utf8
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
