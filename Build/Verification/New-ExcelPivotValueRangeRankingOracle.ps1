<#
.SYNOPSIS
Creates independent Microsoft Excel value-range and top/bottom pivot fixtures.
#>
[CmdletBinding()]
param([string[]] $Kinds = @(), [string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = if ($OutputDirectory) { $OutputDirectory }
    else { Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus' }
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$directory = (Resolve-Path -LiteralPath $directory).Path
if (-not ('OfficeIMOExcelPivotOracleProcess' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OfficeIMOExcelPivotOracleProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
}
$mutex = [Threading.Mutex]::new($false, 'Local\OfficeIMO.Excel.Tests.DesktopCom')
$acquired = $false
$existingExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$application = $null
$workbook = $null
$isolated = $false
try {
    $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5))
    if (-not $acquired) { throw 'Excel COM lock timed out.' }
    $application = New-Object -ComObject Excel.Application
    [uint32]$excelProcessId = 0
    [void][OfficeIMOExcelPivotOracleProcess]::GetWindowThreadProcessId(
        [IntPtr][long]$application.Hwnd, [ref]$excelProcessId)
    if ($excelProcessId -eq 0 -or $existingExcelIds -contains [int]$excelProcessId) {
        throw 'Could not prove isolated Excel instance.'
    }
    $isolated = $true
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $cases = @(
        [pscustomobject]@{ Key = 'between'; Type = 13; First = 20.0; Second = 40.0 },
        [pscustomobject]@{ Key = 'not-between'; Type = 14; First = 20.0; Second = 40.0 },
        [pscustomobject]@{ Key = 'top-count'; Type = 1; First = 1.0; Second = $null },
        [pscustomobject]@{ Key = 'bottom-count'; Type = 2; First = 1.0; Second = $null },
        [pscustomobject]@{ Key = 'top-percent'; Type = 3; First = 40.0; Second = $null },
        [pscustomobject]@{ Key = 'bottom-percent'; Type = 4; First = 40.0; Second = $null },
        [pscustomobject]@{ Key = 'top-sum'; Type = 5; First = 90.0; Second = $null },
        [pscustomobject]@{ Key = 'bottom-sum'; Type = 6; First = 35.0; Second = $null },
        [pscustomobject]@{ Key = 'top-sum-tie'; Type = 5; First = 40.0; Second = $null; Range = 'A4:B6'; Grand = 50.0 },
        [pscustomobject]@{ Key = 'top-percent-tie'; Type = 3; First = 10.0; Second = $null; Range = 'A4:B6'; Grand = 50.0 }
    )
    if (@($Kinds | Where-Object { $_ -notin $cases.Key }).Count -gt 0) {
        throw "Unknown pivot ranking fixture kind: $($Kinds -join ', ')"
    }
    foreach ($case in $cases) {
        if ($Kinds.Count -gt 0 -and $Kinds -notcontains $case.Key) { continue }
        $workbook = $application.Workbooks.Add()
        $source = $workbook.Worksheets.Item(1)
        $source.Name = 'Source'
        $rows = @(
            @('Region', 'Sales'), @('Alpha', 10.0), @('Bravo', 20.0),
            @('Charlie', 30.0), @('Delta', 40.0), @('Echo', 50.0), @('Foxtrot', 50.0)
        )
        $values = New-Object 'object[,]' 7, 2
        for ($row = 0; $row -lt 7; $row++) {
            for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
        }
        $source.Range('A1:B7').Value2 = $values
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
        if ($null -eq $case.Second) {
            [void]$field.PivotFilters.Add2($case.Type, $measure, $case.First)
        } else {
            [void]$field.PivotFilters.Add2($case.Type, $measure, $case.First, $case.Second)
        }
        $lookups = $workbook.Worksheets.Add()
        $lookups.Name = 'Lookups'
        $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
        $names = @('Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot')
        for ($index = 0; $index -lt $names.Count; $index++) {
            $lookups.Cells.Item($index + 2, 2).Formula =
                '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","' + $names[$index] + '")'
        }
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = $pivot.GetPivotData('Metric').Value2
        if ($null -ne $case.Range -and ($range -ne $case.Range -or $grand -ne $case.Grand)) {
            throw "Excel pivot ranking oracle changed: $($case.Key) saved $range and $grand."
        }
        $file = "pivot-value-$($case.Key)-conformance.xlsx"
        $path = Join-Path $directory $file
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
        $workbook = $null
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotValueRangeRankingOracle.ps1'
            file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = 'Source!A1:B7'; filterType = $case.Type
            value1 = $case.First; value2 = $case.Second; outputRange = $range
            grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-value-$($case.Key)-conformance.provenance.json") -Encoding utf8
        [pscustomobject]@{ File = $file; Range = $range; Grand = $grand }
    }
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $application) {
        if ($isolated) { $application.Quit() }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
