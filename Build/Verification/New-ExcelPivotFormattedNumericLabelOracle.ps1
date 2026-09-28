<#
.SYNOPSIS
Creates Microsoft Excel pivot label-filter fixtures with formatted numeric keys.
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
        [pscustomobject]@{ Key = 'contains-comma'; Type = 21; Criterion = ','; Format = '#,##0'; Total = 50.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'equals-grouped'; Type = 15; Criterion = '1,000'; Format = '#,##0'; Total = 20.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'not-equals-grouped'; Type = 16; Criterion = '1,000'; Format = '#,##0'; Total = 40.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'general-contains-one'; Type = 21; Criterion = '1'; Format = 'General'; Total = 30.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'midpoint-positive'; Type = 15; Criterion = '3'; Format = '#,##0'; Values = @('=5/2', '=9/2'); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'midpoint-negative'; Type = 15; Criterion = '-3'; Format = '#,##0'; Values = @('=-5/2', '=9/2'); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'decimal-two'; Type = 15; Criterion = '1.20'; Format = '0.00'; Values = @(1.2, 2.3); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'percent-one'; Type = 15; Criterion = '12.5%'; Format = '0.0%'; Values = @(0.125, 0.25); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'decimal-midpoint'; Type = 15; Criterion = '1.3'; Format = '0.0'; Values = @('=5/4', '=9/4'); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'decimal-three'; Type = 15; Criterion = '1.234'; Format = '0.000'; Values = @(1.2344, 1.2346, 2.5); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'grouped-three-decimal'; Type = 15; Criterion = '1,234.567'; Format = '#,##0.000'; Values = @(1234.5674, 1234.5676, 2000.0); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'range-general-greater'; Type = 23; Criterion = '2'; Format = 'General'; Values = @(1.0, 2.0, 10.0, 20.0); Total = 40.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'range-general-greater-equal'; Type = 24; Criterion = '2'; Format = 'General'; Values = @(1.0, 2.0, 10.0, 20.0); Total = 60.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'range-general-less'; Type = 25; Criterion = '2'; Format = 'General'; Values = @(1.0, 2.0, 10.0, 20.0); Total = 40.0; Range = 'A4:B7' },
        [pscustomobject]@{ Key = 'range-general-less-equal'; Type = 26; Criterion = '2'; Format = 'General'; Values = @(1.0, 2.0, 10.0, 20.0); Total = 60.0; Range = 'A4:B8' },
        [pscustomobject]@{ Key = 'range-general-not-between'; Type = 28; Criterion = '1'; Criterion2 = '2'; Format = 'General'; Values = @(1.0, 2.0, 10.0, 20.0); Total = 40.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'range-grouped-between'; Type = 27; Criterion = '1,000'; Criterion2 = '2,000'; Format = '#,##0'; Values = @(900.0, 1000.0, 1500.0, 2000.0, 3000.0); Total = 90.0; Range = 'A4:B8' },
        [pscustomobject]@{ Key = 'currency-positive'; Type = 15; Criterion = '$1,000.00'; Format = '$#,##0.00'; Values = @(1000, 2000); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'currency-negative'; Type = 15; Criterion = '-$1,000.00'; Format = '$#,##0.00'; Values = @(-1000, 2000); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'grouped-two-decimal'; Type = 15; Criterion = '1,234.50'; Format = '#,##0.00'; Values = @(1234.5, 2000.0); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'percent-two-decimal'; Type = 15; Criterion = '12.50%'; Format = '0.00%'; Values = @(0.125, 0.25); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'currency-zero-decimal'; Type = 15; Criterion = '$1,235'; Format = '$#,##0'; Values = @(1234.5, 2000.0); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'parenthesized-negative'; Type = 15; Criterion = '(1,234.50)'; Format = '#,##0.00;(#,##0.00)'; Values = @(-1234.5, 2000.0); Total = 10.0; Range = 'A4:B6' },
        [pscustomobject]@{ Key = 'duplicate-caption-unfiltered'; Type = $null; Criterion = $null; Format = '0.0'; Values = @(1.21, 1.24, 2.26); Total = 60.0; Range = 'A4:B8' },
        [pscustomobject]@{ Key = 'duplicate-caption-equals'; Type = 15; Criterion = '1.2'; Format = '0.0'; Values = @(1.21, 1.24, 2.26); Total = 30.0; Range = 'A4:B7' }
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
                $source.Cells.Item($index + 2, 1).Value2 = [double]$keys[$index]
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
        if ($null -ne $case.Type) {
            if ($case.Criterion2) { [void]$field.PivotFilters.Add2($case.Type, [Type]::Missing, $case.Criterion, $case.Criterion2) }
            else { [void]$field.PivotFilters.Add2($case.Type, [Type]::Missing, $case.Criterion) }
        }
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
        try { $workbook.Close($false) }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook); $workbook = $null }
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotFormattedNumericLabelOracle.ps1'
            file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = "Source!A1:B$lastSourceRow"; numberFormat = $case.Format
            filterType = $case.Type; criterion = $case.Criterion; criterion2 = $case.Criterion2
            outputRange = $range; grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-label-number-$($case.Key)-conformance.provenance.json") -Encoding utf8
        [pscustomobject]@{ File = $file; Range = $range; Grand = $grand }
    }
} finally {
    try {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
        }
    } finally {
        try {
            if ($null -ne $application) {
                try { if ($isolated) { $application.Quit() } }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) }
            }
        } finally {
            if ($acquired) { $mutex.ReleaseMutex() }
            $mutex.Dispose()
        }
    }
}
