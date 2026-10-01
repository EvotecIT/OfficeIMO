<#
.SYNOPSIS
Creates Excel-produced all-years month and quarter pivot-filter fixtures.
#>
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$targetDirectory = if ($OutputDirectory) { $OutputDirectory }
    else { Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus/CalendarPeriods' }
New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
$targetDirectory = (Resolve-Path -LiteralPath $targetDirectory).Path
if (-not ('OfficeIMOExcelPivotCalendarOracleProcess' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OfficeIMOExcelPivotCalendarOracleProcess {
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
$cases = @(
    for ($month = 1; $month -le 12; $month++) {
        @{ Name = ('month-{0:D2}' -f $month); Type = 56 + $month; Month = $month
            Total = 2 * $month + 12 + $(if ($month -eq 1) { 100 } else { 0 }) }
    }
    for ($quarter = 1; $quarter -le 4; $quarter++) {
        @{ Name = "quarter-$quarter"; Type = 52 + $quarter; Quarter = $quarter
            Total = 18 * $quarter + 30 + $(if ($quarter -eq 1) { 100 } else { 0 }) }
    }
    @{ Name = 'month-01-1904'; Type = 57; Month = 1; Date1904 = $true; Total = 114 }
    @{ Name = 'quarter-1-1904'; Type = 53; Quarter = 1; Date1904 = $true; Total = 148 }
)
$results = @()
try {
    $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5))
    if (-not $acquired) { throw 'Excel COM lock timed out.' }
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcessId = 0
    [void][OfficeIMOExcelPivotCalendarOracleProcess]::GetWindowThreadProcessId(
        [IntPtr][long]$excel.Hwnd, [ref]$excelProcessId)
    if ($excelProcessId -eq 0 -or $existingExcelIds -contains [int]$excelProcessId) {
        throw 'Could not prove isolated Excel instance.'
    }
    $isolated = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    foreach ($case in $cases) {
        $workbook = $excel.Workbooks.Add()
        try {
            if ($case.Date1904) { $workbook.Date1904 = $true }
            $source = $workbook.Worksheets.Item(1)
            $source.Name = 'Source'
            $source.Range('A1').Value2 = 'OrderDate'
            $source.Range('B1').Value2 = 'Sales'
            for ($index = 0; $index -lt 24; $index++) {
                $row = $index + 2
                $year = if ($index -lt 12) { 2024 } else { 2025 }
                $month = $index % 12 + 1
                $source.Range("A$row").Value = [datetime]::new($year, $month, 1)
                $source.Range("B$row").Value2 = [double]($index + 1)
            }
            $source.Range('A26').Value = [datetime]'2025-01-15T08:30:00'
            $source.Range('B26').Value2 = 100.0
            $source.Range('B27').Value2 = 1000.0
            [void]$source.Columns.Item(1).AutoFit()
            $view = $workbook.Worksheets.Add()
            $view.Name = 'Grouped'
            $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R27C2", 6)
            $pivot = $cache.CreatePivotTable($view.Range('A4'), 'DatePivot')
            $field = $pivot.PivotFields('OrderDate')
            $field.Orientation = 1
            $field.Position = 1
            [void]$pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
            $pivot.RowAxisLayout(1)
            [void]$pivot.RefreshTable()
            [void]$field.PivotFilters.Add2($case.Type)
            $siblingRange = $null
            $siblingRows = @()
            if ($case.Name -eq 'month-01') {
                $sibling = $workbook.Worksheets.Add()
                $sibling.Name = 'AllDates'
                $siblingPivot = $cache.CreatePivotTable($sibling.Range('A4'), 'AllDatesPivot')
                $siblingField = $siblingPivot.PivotFields('OrderDate')
                $siblingField.Orientation = 1
                $siblingField.Position = 1
                [void]$siblingPivot.AddDataField($siblingPivot.PivotFields('Sales'), 'Metric', -4157)
                $siblingPivot.RowAxisLayout(1)
                [void]$siblingPivot.RefreshTable()
                # Keep a deliberately different manual view order on the same cache.
                $siblingField.PivotItems().Item(25).Position = 1
                [void]$sibling.Columns.Item(1).AutoFit()
                $siblingRange = $siblingPivot.TableRange1.Address($false, $false)
                if ([double]$siblingPivot.GetPivotData('Metric').Value2 -ne 1400.0) {
                    throw 'Unexpected unfiltered shared-cache total.'
                }
                for ($row = 4; $row -lt 4 + [int]$siblingPivot.TableRange1.Rows.Count; $row++) {
                    $siblingRows += [ordered]@{ label = [string]$sibling.Cells.Item($row, 1).Text
                        metric = [string]$sibling.Cells.Item($row, 2).Text }
                }
            }
            $excel.CalculateFullRebuild()
            [void]$view.Columns.Item(1).AutoFit()
            $range = $pivot.TableRange1.Address($false, $false)
            $grand = [double]$pivot.GetPivotData('Metric').Value2
            if ($grand -ne $case.Total) {
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
            name = $case.Name; filterType = $case.Type
            month = if ($case.Month) { [int]$case.Month } else { $null }
            quarter = if ($case.Quarter) { [int]$case.Quarter } else { $null }
            date1904 = [bool]$case.Date1904
            sourceRange = 'Source!A1:B27'; sourceDateStyle = 'Excel automatic DateTime'
            outputRange = $range; grandTotal = $grand
            siblingRange = $siblingRange; siblingRows = $siblingRows
            rows = $viewRows; file = [IO.Path]::GetFileName($path)
            sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    $manifest = [ordered]@{
        producer = 'Microsoft Excel'; version = $excel.Version; build = $excel.Build
        producerCulture = [Globalization.CultureInfo]::CurrentCulture.Name
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotCalendarPeriodOracle.ps1'
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
