<#
.SYNOPSIS
Creates Excel-produced fixed-date pivot filter fixtures with saved-view provenance.
#>
[CmdletBinding()]
param([string] $OutputDirectory, [string[]] $Kinds = @())
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$targetDirectory = if ($OutputDirectory) { $OutputDirectory }
    else { Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus/FixedDateFilters' }
New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
$targetDirectory = (Resolve-Path -LiteralPath $targetDirectory).Path
if (-not ('OfficeIMOExcelPivotDateOracleProcess' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OfficeIMOExcelPivotDateOracleProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
}
$mutex = [Threading.Mutex]::new($false, 'Local\OfficeIMO.Excel.Tests.DesktopCom')
$acquired = $false
$existingExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$excel = $null
$workbook = $null
$isolated = $false
$first = [datetime]'2025-01-15'
$middle = [datetime]'2025-03-20'
$last = [datetime]'2026-01-15'
$cases = @(
    @{ Name = 'equal'; Type = 29; First = $first; Total = 15.0 },
    @{ Name = 'not-equal'; Type = 30; First = $first; Total = 50.0 },
    @{ Name = 'before'; Type = 31; First = $last; Total = 35.0 },
    @{ Name = 'before-equal'; Type = 32; First = $last; Total = 65.0 },
    @{ Name = 'after'; Type = 33; First = $first; Total = 50.0 },
    @{ Name = 'after-equal'; Type = 34; First = $first; Total = 65.0 },
    @{ Name = 'between'; Type = 35; First = $first; Second = $middle; Total = 35.0 },
    @{ Name = 'not-between'; Type = 36; First = $first; Second = $middle; Total = 30.0 },
    @{ Name = 'equal-times'; Type = 29; First = $first; Times = $true; Total = 10.0 },
    @{ Name = 'between-times'; Type = 35; First = $first; Second = $middle; Times = $true; Total = 15.0 },
    @{ Name = 'after-times'; Type = 33; First = $first; Times = $true; Total = 55.0 },
    @{ Name = 'equal-1904'; Type = 29; First = $first; Date1904 = $true; Total = 15.0 },
    @{ Name = 'between-1904'; Type = 35; First = $first; Second = $middle; Date1904 = $true; Total = 35.0 },
    @{ Name = 'equal-blank'; Type = 29; First = $first; BlankSource = $true },
    @{ Name = 'not-equal-blank'; Type = 30; First = $first; BlankSource = $true }
)
if (@($Kinds | Where-Object { $_ -notin $cases.Name }).Count -gt 0) {
    throw "Unknown date-filter fixture kind: $($Kinds -join ', ')"
}
$results = @()
try {
    $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5))
    if (-not $acquired) { throw 'Excel COM lock timed out.' }
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcessId = 0
    [void][OfficeIMOExcelPivotDateOracleProcess]::GetWindowThreadProcessId(
        [IntPtr][long]$excel.Hwnd, [ref]$excelProcessId)
    if ($excelProcessId -eq 0 -or $existingExcelIds -contains [int]$excelProcessId) {
        throw 'Could not prove isolated Excel instance.'
    }
    $isolated = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    foreach ($case in $cases) {
        if ($Kinds.Count -gt 0 -and $Kinds -notcontains $case.Name) { continue }
        $workbook = $excel.Workbooks.Add()
        try {
            if ($case.Date1904) { $workbook.Date1904 = $true }
            $source = $workbook.Worksheets.Item(1)
            $source.Name = 'Source'
            $source.Range('A1').Value2 = 'OrderDate'
            $source.Range('B1').Value2 = 'Sales'
            $rows = if ($case.BlankSource) {
                @(
                    @($first, 10.0), @($null, 5.0),
                    @($middle, 20.0), @($last, 30.0)
                )
            } elseif ($case.Times) {
                @(
                    @($first, 10.0), @($first.AddHours(8).AddMinutes(30), 5.0),
                    @($middle.AddHours(12), 20.0), @($last, 30.0)
                )
            } else {
                @(
                    @($first, 10.0), @($first, 5.0),
                    @($middle, 20.0), @($last, 30.0)
                )
            }
            for ($index = 0; $index -lt $rows.Count; $index++) {
                $row = $index + 2
                if ($null -ne $rows[$index][0]) {
                    $source.Range("A$row").Value = [datetime]$rows[$index][0]
                }
                $source.Range("B$row").Value2 = [double]$rows[$index][1]
            }
            [void]$source.Columns.Item(1).AutoFit()
            $view = $workbook.Worksheets.Add()
            $view.Name = 'Grouped'
            $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R5C2", 6)
            $pivot = $cache.CreatePivotTable($view.Range('A4'), 'DatePivot')
            $field = $pivot.PivotFields('OrderDate')
            $field.Orientation = 1
            $field.Position = 1
            [void]$pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
            $pivot.RowAxisLayout(1)
            [void]$pivot.RefreshTable()
            if ($case.Second) {
                [void]$field.PivotFilters.Add2($case.Type, [Type]::Missing, $case.First, $case.Second)
            } else {
                [void]$field.PivotFilters.Add2($case.Type, [Type]::Missing, $case.First)
            }
            $excel.CalculateFullRebuild()
            [void]$view.Columns.Item(1).AutoFit()
            $range = $pivot.TableRange1.Address($false, $false)
            $grand = [double]$pivot.GetPivotData('Metric').Value2
            if ($case.ContainsKey('Total') -and $grand -ne $case.Total) {
                throw "Unexpected Excel total for $($case.Name): $grand"
            }
            $viewRows = @()
            for ($row = 4; $row -lt 4 + [int]$pivot.TableRange1.Rows.Count; $row++) {
                $viewRows += [ordered]@{ label = [string]$view.Cells.Item($row, 1).Text
                    metric = [string]$view.Cells.Item($row, 2).Text }
            }
            $path = Join-Path $targetDirectory ("$($case.Name).xlsx")
            $workbook.SaveAs($path, 51)
        } finally {
            try { $workbook.Close($false) }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook); $workbook = $null }
        }
        $results += [ordered]@{
            name = $case.Name; filterType = $case.Type; timedSource = [bool]$case.Times
            blankSource = [bool]$case.BlankSource; date1904 = [bool]$case.Date1904
            first = $case.First.ToString('o'); second = if ($case.Second) { $case.Second.ToString('o') } else { $null }
            sourceRange = 'Source!A1:B5'; sourceDateStyle = 'Excel automatic DateTime'
            outputRange = $range; grandTotal = $grand
            rows = $viewRows; file = [IO.Path]::GetFileName($path)
            sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    $manifest = [ordered]@{
        producer = 'Microsoft Excel'; version = $excel.Version; build = $excel.Build
        producerCulture = [Globalization.CultureInfo]::CurrentCulture.Name
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotFixedDateFilterOracle.ps1'
        cases = $results
    }
    [IO.File]::WriteAllText((Join-Path $targetDirectory 'provenance.json'),
        ($manifest | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    $results | ForEach-Object { [pscustomobject]$_ } |
        Select-Object name,outputRange,grandTotal | Format-Table -AutoSize | Out-String | Write-Output
} finally {
    try {
        if ($workbook -ne $null) {
            try { $workbook.Close($false) }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
        }
    } finally {
        try {
            if ($excel -ne $null) {
                try { if ($isolated) { $excel.Quit() } }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
            }
        } finally {
            if ($acquired) { $mutex.ReleaseMutex() }
            $mutex.Dispose()
        }
    }
}
