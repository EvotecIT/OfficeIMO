<#
.SYNOPSIS
Creates Excel-produced relative-date pivot-filter fixtures.
#>
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$targetDirectory = if ($OutputDirectory) { $OutputDirectory }
    else { Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus/RelativeDates' }
New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
$targetDirectory = (Resolve-Path -LiteralPath $targetDirectory).Path
if (-not ('OfficeIMOExcelPivotRelativeOracleProcess' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OfficeIMOExcelPivotRelativeOracleProcess {
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
$today = (Get-Date).Date
$weekStart = $today.AddDays(-[int]$today.DayOfWeek)
$monthStart = [datetime]::new($today.Year, $today.Month, 1)
$quarterStart = [datetime]::new($today.Year, (3 * [math]::Floor(($today.Month - 1) / 3)) + 1, 1)
$yearStart = [datetime]::new($today.Year, 1, 1)
$dates = @(
    $today.AddDays(-1); $today; $today.AddDays(1); $today.AddHours(8)
    $weekStart.AddDays(-8); $weekStart.AddDays(-7); $weekStart.AddDays(-1)
    $weekStart; $weekStart.AddDays(6); $weekStart.AddDays(7)
    $monthStart.AddMonths(-1); $monthStart.AddDays(-1); $monthStart
    $monthStart.AddMonths(1).AddDays(-1); $monthStart.AddMonths(1)
    $quarterStart.AddMonths(-3); $quarterStart.AddDays(-1); $quarterStart
    $quarterStart.AddMonths(3).AddDays(-1); $quarterStart.AddMonths(3)
    $yearStart.AddYears(-1); $yearStart.AddDays(-1); $yearStart
    $yearStart.AddYears(1).AddDays(-1); $yearStart.AddYears(1)
) | Sort-Object -Unique
$cases = @(
    @{ Name = 'yesterday'; Type = 39 }; @{ Name = 'today'; Type = 38 }
    @{ Name = 'tomorrow'; Type = 37 }; @{ Name = 'last-week'; Type = 42 }
    @{ Name = 'this-week'; Type = 41 }; @{ Name = 'next-week'; Type = 40 }
    @{ Name = 'last-month'; Type = 45 }; @{ Name = 'this-month'; Type = 44 }
    @{ Name = 'next-month'; Type = 43 }; @{ Name = 'last-quarter'; Type = 48 }
    @{ Name = 'this-quarter'; Type = 47 }; @{ Name = 'next-quarter'; Type = 46 }
    @{ Name = 'last-year'; Type = 51 }; @{ Name = 'this-year'; Type = 50 }
    @{ Name = 'next-year'; Type = 49 }; @{ Name = 'year-to-date'; Type = 52 }
    @{ Name = 'this-month-1904'; Type = 44; Date1904 = $true }
)
$results = @()
try {
    $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5))
    if (-not $acquired) { throw 'Excel COM lock timed out.' }
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelProcessId = 0
    [void][OfficeIMOExcelPivotRelativeOracleProcess]::GetWindowThreadProcessId(
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
            for ($index = 0; $index -lt $dates.Count; $index++) {
                $row = $index + 2
                $source.Range("A$row").Value = [datetime]$dates[$index]
                $source.Range("B$row").Value2 = [double]($index + 1)
            }
            $blankRow = $dates.Count + 2
            $source.Range("B$blankRow").Value2 = 1000.0
            [void]$source.Columns.Item(1).AutoFit()
            $view = $workbook.Worksheets.Add()
            $view.Name = 'Grouped'
            $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R${blankRow}C2", 6)
            $pivot = $cache.CreatePivotTable($view.Range('A4'), 'DatePivot')
            $field = $pivot.PivotFields('OrderDate')
            $field.Orientation = 1
            $field.Position = 1
            [void]$pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
            $pivot.RowAxisLayout(1)
            [void]$pivot.RefreshTable()
            [void]$field.PivotFilters.Add2($case.Type)
            $siblingRange = $null
            if ($case.Name -eq 'today') {
                $sibling = $workbook.Worksheets.Add()
                $sibling.Name = 'AllDates'
                $siblingPivot = $cache.CreatePivotTable($sibling.Range('A4'), 'AllDatesPivot')
                $siblingField = $siblingPivot.PivotFields('OrderDate')
                $siblingField.Orientation = 1
                $siblingField.Position = 1
                [void]$siblingPivot.AddDataField($siblingPivot.PivotFields('Sales'), 'Metric', -4157)
                $siblingPivot.RowAxisLayout(1)
                [void]$siblingPivot.RefreshTable()
                $siblingField.PivotItems().Item($dates.Count).Position = 1
                [void]$sibling.Columns.Item(1).AutoFit()
                $siblingRange = $siblingPivot.TableRange1.Address($false, $false)
                $expectedAllDates = ($dates.Count * ($dates.Count + 1) / 2) + 1000
                if ([double]$siblingPivot.GetPivotData('Metric').Value2 -ne $expectedAllDates)
                    { throw 'Unexpected unfiltered shared-cache total.' }
            }
            $excel.CalculateFullRebuild()
            [void]$view.Columns.Item(1).AutoFit()
            $range = $pivot.TableRange1.Address($false, $false)
            $grand = [double]$pivot.GetPivotData('Metric').Value2
            if ($grand -le 0 -or $grand -ge 1000) { throw "Unexpected Excel total for $($case.Name): $grand" }
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
            date1904 = [bool]$case.Date1904
            sourceRange = "Source!A1:B$blankRow"; sourceDateStyle = 'Excel automatic DateTime'
            outputRange = $range; grandTotal = $grand
            siblingRange = $siblingRange
            rows = $viewRows; file = [IO.Path]::GetFileName($path)
            sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    $manifest = [ordered]@{
        producer = 'Microsoft Excel'; version = $excel.Version; build = $excel.Build
        producerCulture = [Globalization.CultureInfo]::CurrentCulture.Name
        referenceDate = $today.ToString('yyyy-MM-dd')
        sourceDates = @($dates | ForEach-Object { ([datetime]$_).ToString('o') })
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotRelativeDateOracle.ps1'
        cases = $results
    }
    [IO.File]::WriteAllText((Join-Path $targetDirectory 'provenance.json'),
        ($manifest | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    if ((Get-Date).Date -ne $today) { throw 'The local date changed while producing relative-date fixtures.' }
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
