<#
.SYNOPSIS
Creates Microsoft Excel pivot value-filter fixtures with two row or column fields.
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
    # Type IDs are Excel XlPivotFilterType values:
    # https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlpivotfiltertype
    $cases = @(
        [pscustomobject]@{ Key = 'outer'; Field = 'Region'; Type = 9; Threshold = 62.0; Axis = 'Row'; Range = 'A4:C8'; Grand = 65.0 },
        [pscustomobject]@{ Key = 'inner'; Field = 'Product'; Type = 9; Threshold = 30.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 150.0 },
        [pscustomobject]@{ Key = 'inner-top1'; Field = 'Product'; Type = 1; Threshold = 1.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 150.0 },
        [pscustomobject]@{ Key = 'column-inner'; Field = 'Product'; Type = 9; Threshold = 30.0; Axis = 'Column'; Range = 'A4:H7'; Grand = 150.0 },
        [pscustomobject]@{ Key = 'inner-less'; Field = 'Product'; Type = 11; Threshold = 30.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 35.0 },
        [pscustomobject]@{ Key = 'inner-between'; Field = 'Product'; Type = 13; Threshold = 15.0; Threshold2 = 45.0; Axis = 'Row'; Range = 'A4:C8'; Grand = 60.0 },
        [pscustomobject]@{ Key = 'inner-notbetween'; Field = 'Product'; Type = 14; Threshold = 15.0; Threshold2 = 45.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 125.0 },
        [pscustomobject]@{ Key = 'inner-bottom1'; Field = 'Product'; Type = 2; Threshold = 1.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 35.0 },
        [pscustomobject]@{ Key = 'inner-top50pct'; Field = 'Product'; Type = 3; Threshold = 50.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 150.0 },
        [pscustomobject]@{ Key = 'inner-topsum30'; Field = 'Product'; Type = 5; Threshold = 30.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 150.0 },
        [pscustomobject]@{ Key = 'outer-bottom1'; Field = 'Region'; Type = 2; Threshold = 1.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'mixed-column'; Field = 'Product'; Type = 9; Threshold = 100.0; Axis = 'Mixed'; Range = 'A4:C9'; Grand = 130.0 },
        [pscustomobject]@{ Key = 'mixed-row'; Field = 'Region'; Type = 9; Threshold = 62.0; Axis = 'Mixed'; Range = 'A4:D7'; Grand = 65.0 },
        [pscustomobject]@{ Key = 'mixed-column-top1'; Field = 'Product'; Type = 1; Threshold = 1.0; Axis = 'Mixed'; Range = 'A4:C9'; Grand = 130.0 },
        [pscustomobject]@{ Key = 'inner-equal'; Field = 'Product'; Type = 7; Threshold = 40.0; Axis = 'Row'; Range = 'A4:C7'; Grand = 40.0 },
        [pscustomobject]@{ Key = 'inner-notequal'; Field = 'Product'; Type = 8; Threshold = 40.0; Axis = 'Row'; Range = 'A4:C13'; Grand = 145.0 },
        [pscustomobject]@{ Key = 'inner-greater-equal'; Field = 'Product'; Type = 10; Threshold = 40.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 150.0 },
        [pscustomobject]@{ Key = 'inner-less-equal'; Field = 'Product'; Type = 12; Threshold = 20.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 35.0 },
        [pscustomobject]@{ Key = 'inner-bottom25pct'; Field = 'Product'; Type = 4; Threshold = 25.0; Axis = 'Row'; Range = 'A4:C13'; Grand = 145.0 },
        [pscustomobject]@{ Key = 'inner-bottomsum15'; Field = 'Product'; Type = 6; Threshold = 15.0; Axis = 'Row'; Range = 'A4:C13'; Grand = 145.0 },
        [pscustomobject]@{ Key = 'outer-top2'; Field = 'Region'; Type = 1; Threshold = 2.0; Axis = 'Row'; Range = 'A4:C14'; Grand = 185.0 },
        [pscustomobject]@{ Key = 'outer-equal'; Field = 'Region'; Type = 7; Threshold = 60.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'outer-notequal'; Field = 'Region'; Type = 8; Threshold = 60.0; Axis = 'Row'; Range = 'A4:C8'; Grand = 65.0 },
        [pscustomobject]@{ Key = 'outer-greater-equal'; Field = 'Region'; Type = 10; Threshold = 65.0; Axis = 'Row'; Range = 'A4:C8'; Grand = 65.0 },
        [pscustomobject]@{ Key = 'outer-less'; Field = 'Region'; Type = 11; Threshold = 65.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'outer-less-equal'; Field = 'Region'; Type = 12; Threshold = 60.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'outer-between'; Field = 'Region'; Type = 13; Threshold = 61.0; Threshold2 = 65.0; Axis = 'Row'; Range = 'A4:C8'; Grand = 65.0 },
        [pscustomobject]@{ Key = 'outer-notbetween'; Field = 'Region'; Type = 14; Threshold = 61.0; Threshold2 = 65.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'inner-wide-top2'; Field = 'Product'; Type = 1; Threshold = 2.0; Axis = 'Row'; Profile = 'WideSigned'; Range = 'A4:C14'; Grand = 210.0 },
        [pscustomobject]@{ Key = 'inner-wide-bottom2'; Field = 'Product'; Type = 2; Threshold = 2.0; Axis = 'Row'; Profile = 'WideSigned'; Range = 'A4:C14'; Grand = 45.0 },
        [pscustomobject]@{ Key = 'inner-wide-top50pct'; Field = 'Product'; Type = 3; Threshold = 50.0; Axis = 'Row'; Profile = 'WideSigned'; Range = 'A4:C11'; Grand = 150.0 },
        [pscustomobject]@{ Key = 'inner-wide-bottomsum15'; Field = 'Product'; Type = 6; Threshold = 15.0; Axis = 'Row'; Profile = 'WideSigned'; Range = 'A4:C15'; Grand = 95.0 },
        [pscustomobject]@{ Key = 'inner-wide-error-top2'; Field = 'Product'; Type = 1; Threshold = 2.0; Axis = 'Row'; Profile = 'WideMixedError'; Range = 'A4:C14'; Grand = 210.0 },
        [pscustomobject]@{ Key = 'outer-wide-top50pct'; Field = 'Region'; Type = 3; Threshold = 50.0; Axis = 'Row'; Profile = 'WideSigned'; Range = 'A4:C13'; Grand = 155.0 },
        [pscustomobject]@{ Key = 'outer-wide-bottom50pct'; Field = 'Region'; Type = 4; Threshold = 50.0; Axis = 'Row'; Profile = 'WideSigned'; Range = 'A4:C13'; Grand = 100.0 },
        [pscustomobject]@{ Key = 'outer-wide-topsum50'; Field = 'Region'; Type = 5; Threshold = 50.0; Axis = 'Row'; Profile = 'WideSigned'; Range = 'A4:C9'; Grand = 95.0 },
        [pscustomobject]@{ Key = 'outer-wide-bottomsum50'; Field = 'Region'; Type = 6; Threshold = 50.0; Axis = 'Row'; Profile = 'WideSigned'; Range = 'A4:C13'; Grand = 100.0 },
        [pscustomobject]@{ Key = 'outer-negative-top40pct'; Field = 'Region'; Type = 3; Threshold = 40.0; Axis = 'Row'; Profile = 'NegativeParents'; Range = 'A4:C13'; Grand = -90.0 },
        [pscustomobject]@{ Key = 'outer-negative-bottom40pct'; Field = 'Region'; Type = 4; Threshold = 40.0; Axis = 'Row'; Profile = 'NegativeParents'; Range = 'A4:C9'; Grand = -60.0 },
        [pscustomobject]@{ Key = 'outer-negative-topsum30'; Field = 'Region'; Type = 5; Threshold = 30.0; Axis = 'Row'; Profile = 'NegativeParents'; Range = 'A4:C17'; Grand = -150.0 },
        [pscustomobject]@{ Key = 'outer-negative-bottomsum30'; Field = 'Region'; Type = 6; Threshold = 30.0; Axis = 'Row'; Profile = 'NegativeParents'; Range = 'A4:C17'; Grand = -150.0 },
        [pscustomobject]@{ Key = 'outer-zero-bottom1'; Field = 'Region'; Type = 2; Threshold = 1.0; Axis = 'Row'; Profile = 'ZeroParents'; Range = 'A4:C13'; Grand = 0.0 },
        [pscustomobject]@{ Key = 'outer-zero-top50pct'; Field = 'Region'; Type = 3; Threshold = 50.0; Axis = 'Row'; Profile = 'ZeroParents'; Range = 'A4:C9'; Grand = 40.0 },
        [pscustomobject]@{ Key = 'outer-zero-bottom50pct'; Field = 'Region'; Type = 4; Threshold = 50.0; Axis = 'Row'; Profile = 'ZeroParents'; Range = 'A4:C17'; Grand = 40.0 },
        [pscustomobject]@{ Key = 'inner-errorparent-top1'; Field = 'Product'; Type = 1; Threshold = 1.0; Axis = 'Row'; Profile = 'ErrorParent'; Range = 'A4:C11'; Grand = '#DIV/0!' },
        [pscustomobject]@{ Key = 'inner-errorparent-bottom1'; Field = 'Product'; Type = 2; Threshold = 1.0; Axis = 'Row'; Profile = 'ErrorParent'; Range = 'A4:C11'; Grand = '#N/A' },
        [pscustomobject]@{ Key = 'mixed-column-less'; Field = 'Product'; Type = 11; Threshold = 100.0; Axis = 'Mixed'; Range = 'A4:C9'; Grand = 55.0 },
        [pscustomobject]@{ Key = 'mixed-column-equal'; Field = 'Product'; Type = 7; Threshold = 55.0; Axis = 'Mixed'; Range = 'A4:C9'; Grand = 55.0 },
        [pscustomobject]@{ Key = 'mixed-column-notequal'; Field = 'Product'; Type = 8; Threshold = 55.0; Axis = 'Mixed'; Range = 'A4:C9'; Grand = 130.0 },
        [pscustomobject]@{ Key = 'mixed-column-between'; Field = 'Product'; Type = 13; Threshold = 50.0; Threshold2 = 60.0; Axis = 'Mixed'; Range = 'A4:C9'; Grand = 55.0 },
        [pscustomobject]@{ Key = 'mixed-column-bottom1'; Field = 'Product'; Type = 2; Threshold = 1.0; Axis = 'Mixed'; Range = 'A4:C9'; Grand = 55.0 },
        [pscustomobject]@{ Key = 'mixed-column-top2'; Field = 'Product'; Type = 1; Threshold = 2.0; Axis = 'Mixed'; Range = 'A4:D9'; Grand = 185.0 },
        [pscustomobject]@{ Key = 'mixed-column-top50pct'; Field = 'Product'; Type = 3; Threshold = 50.0; Axis = 'Mixed'; Range = 'A4:C9'; Grand = 130.0 },
        [pscustomobject]@{ Key = 'mixed-column-bottomsum60'; Field = 'Product'; Type = 6; Threshold = 60.0; Axis = 'Mixed'; Range = 'A4:D9'; Grand = 185.0 },
        [pscustomobject]@{ Key = 'mixed-row-less'; Field = 'Region'; Type = 11; Threshold = 62.0; Axis = 'Mixed'; Range = 'A4:D8'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'mixed-row-notequal'; Field = 'Region'; Type = 8; Threshold = 60.0; Axis = 'Mixed'; Range = 'A4:D7'; Grand = 65.0 },
        [pscustomobject]@{ Key = 'mixed-row-between'; Field = 'Region'; Type = 13; Threshold = 61.0; Threshold2 = 65.0; Axis = 'Mixed'; Range = 'A4:D7'; Grand = 65.0 },
        [pscustomobject]@{ Key = 'mixed-row-notbetween'; Field = 'Region'; Type = 14; Threshold = 61.0; Threshold2 = 65.0; Axis = 'Mixed'; Range = 'A4:D8'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'mixed-row-bottom1'; Field = 'Region'; Type = 2; Threshold = 1.0; Axis = 'Mixed'; Range = 'A4:D8'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'mixed-row-top2'; Field = 'Region'; Type = 1; Threshold = 2.0; Axis = 'Mixed'; Range = 'A4:D9'; Grand = 185.0 },
        [pscustomobject]@{ Key = 'mixed-row-top25pct'; Field = 'Region'; Type = 3; Threshold = 25.0; Axis = 'Mixed'; Range = 'A4:D7'; Grand = 65.0 },
        [pscustomobject]@{ Key = 'mixed-row-bottomsum50'; Field = 'Region'; Type = 6; Threshold = 50.0; Axis = 'Mixed'; Range = 'A4:D7'; Grand = 60.0 },
        [pscustomobject]@{ Key = 'mixed-row-bottomsum70'; Field = 'Region'; Type = 6; Threshold = 70.0; Axis = 'Mixed'; Range = 'A4:D8'; Grand = 120.0 },
        [pscustomobject]@{ Key = 'outer-tie-bottomsum50'; Field = 'Region'; Type = 6; Threshold = 50.0; Axis = 'Row'; Range = 'A4:C8'; Grand = 60.0 },
        [pscustomobject]@{ Key = 'outer-tie-top50pct'; Field = 'Region'; Type = 3; Threshold = 50.0; Axis = 'Row'; Range = 'A4:C11'; Grand = 125.0 },
        [pscustomobject]@{ Key = 'outer-tie-bottom25pct'; Field = 'Region'; Type = 4; Threshold = 25.0; Axis = 'Row'; Range = 'A4:C8'; Grand = 60.0 },
        [pscustomobject]@{ Key = 'outer-tie-bottomsum50-reversed'; Field = 'Region'; Type = 6; Threshold = 50.0; Axis = 'Row'; Profile = 'ReverseParents'; Range = 'A4:C8'; Grand = 60.0 },
        [pscustomobject]@{ Key = 'outer-tie-bottomsum50-three'; Field = 'Region'; Type = 6; Threshold = 50.0; Axis = 'Row'; Profile = 'ThreeTies'; Range = 'A4:C8'; Grand = 60.0 }
    )
    if (@($Kinds | Where-Object { $_ -notin $cases.Key }).Count -gt 0) {
        throw "Unknown pivot value fixture kind: $($Kinds -join ', ')"
    }
    foreach ($case in $cases) {
        if ($Kinds.Count -gt 0 -and $Kinds -notcontains $case.Key) { continue }
        $workbook = $application.Workbooks.Add()
        $source = $workbook.Worksheets.Item(1)
        $source.Name = 'Source'
        @('Region', 'Product', 'Sales') | ForEach-Object -Begin { $column = 1 } -Process {
            $source.Cells.Item(1, $column).Value2 = $_
            $column++
        }
        $rows = if ($case.Profile -eq 'WideSigned' -or $case.Profile -eq 'WideMixedError') {
            @(
                @('East', 'A', 10.0), @('East', 'B', 50.0),
                @('East', 'C', $(if ($case.Profile -eq 'WideMixedError') { '=NA()' } else { -20.0 })),
                @('West', 'A', 40.0), @('West', 'B', 20.0), @('West', 'C', 0.0),
                @('South', 'A', 5.0), @('South', 'B', 60.0), @('South', 'C', 30.0)
            )
        } elseif ($case.Profile -eq 'NegativeParents') {
            @(
                @('East', 'A', -30.0), @('East', 'B', -30.0), @('East', 'C', 0.0),
                @('West', 'A', -30.0), @('West', 'B', -20.0), @('West', 'C', 0.0),
                @('South', 'A', -20.0), @('South', 'B', -20.0), @('South', 'C', 0.0)
            )
        } elseif ($case.Profile -eq 'ZeroParents') {
            @(
                @('East', 'A', -10.0), @('East', 'B', 10.0), @('East', 'C', 0.0),
                @('West', 'A', -5.0), @('West', 'B', 5.0), @('West', 'C', 0.0),
                @('South', 'A', 20.0), @('South', 'B', 20.0), @('South', 'C', 0.0)
            )
        } elseif ($case.Profile -eq 'ErrorParent') {
            @(
                @('East', 'A', '=NA()'), @('East', 'B', '=1/0'), @('East', 'C', '=VALUE("bad")'),
                @('West', 'A', 40.0), @('West', 'B', 20.0), @('West', 'C', 0.0),
                @('South', 'A', 5.0), @('South', 'B', 60.0), @('South', 'C', 30.0)
            )
        } elseif ($case.Profile -eq 'ReverseParents') {
            @(
                @('South', 'A', 5.0), @('South', 'B', 60.0),
                @('West', 'A', 40.0), @('West', 'B', 20.0),
                @('East', 'A', 10.0), @('East', 'B', 50.0)
            )
        } elseif ($case.Profile -eq 'ThreeTies') {
            @(
                @('East', 'A', 10.0), @('East', 'B', 50.0),
                @('West', 'A', 40.0), @('West', 'B', 20.0),
                @('South', 'A', 5.0), @('South', 'B', 55.0)
            )
        } else {
            @(
                @('East', 'A', 10.0), @('East', 'B', 50.0),
                @('West', 'A', 40.0), @('West', 'B', 20.0),
                @('South', 'A', 5.0), @('South', 'B', 60.0)
            )
        }
        for ($index = 0; $index -lt $rows.Count; $index++) {
            for ($column = 0; $column -lt 3; $column++) {
                if ($column -eq 2) {
                    if ($rows[$index][$column] -is [string] -and $rows[$index][$column].StartsWith('=')) {
                        $source.Cells.Item($index + 2, $column + 1).Formula = $rows[$index][$column]
                    } else {
                        $source.Cells.Item($index + 2, $column + 1).Value2 = [double]$rows[$index][$column]
                    }
                } else {
                    $source.Cells.Item($index + 2, $column + 1).Value2 = [string]$rows[$index][$column]
                }
            }
        }
        if ($case.Profile -eq 'WideMixedError' -or $case.Profile -eq 'ErrorParent') {
            $application.CalculateFullRebuild()
        }
        $view = $workbook.Worksheets.Add()
        $view.Name = 'Grouped'
        $sourceRange = "'Source'!R1C1:R$($rows.Count + 1)C3"
        $cache = $workbook.PivotCaches().Create(1, $sourceRange, 6)
        $pivot = $cache.CreatePivotTable($view.Range('A4'), 'ValuePivot')
        $region = $pivot.PivotFields('Region')
        $region.Orientation = if ($case.Axis -eq 'Column') { 2 } else { 1 }
        $region.Position = 1
        $product = $pivot.PivotFields('Product')
        $product.Orientation = if ($case.Axis -eq 'Column' -or $case.Axis -eq 'Mixed') { 2 } else { 1 }
        $product.Position = if ($case.Axis -eq 'Mixed') { 1 } else { 2 }
        $metric = $pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
        if ($case.Axis -ne 'Column') { $pivot.RowAxisLayout(1) }
        [void]$pivot.RefreshTable()
        $target = $pivot.PivotFields($case.Field)
        if ($null -ne $case.Threshold2) {
            [void]$target.PivotFilters.Add2($case.Type, $metric, $case.Threshold, $case.Threshold2)
        } else {
            [void]$target.PivotFilters.Add2($case.Type, $metric, $case.Threshold)
        }
        $lookups = $workbook.Worksheets.Add()
        $lookups.Name = 'Lookups'
        $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
        $lookupRows = @($rows | ForEach-Object { ,@([string]$_[0], [string]$_[1]) })
        for ($index = 0; $index -lt $lookupRows.Count; $index++) {
            $regionName = $lookupRows[$index][0]
            $productName = $lookupRows[$index][1]
            $lookups.Cells.Item($index + 2, 2).Formula =
                '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","' + $regionName +
                '","Product","' + $productName + '")'
        }
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = if ($case.Grand -is [string]) {
            [string]$pivot.GetPivotData('Metric').Text
        } else {
            [double]$pivot.GetPivotData('Metric').Value2
        }
        if ($range -ne $case.Range -or $grand -ne $case.Grand) {
            throw "Excel pivot oracle changed: $($case.Key) saved $range and $grand."
        }
        $file = "pivot-value-multifield-$($case.Key)-conformance.xlsx"
        $path = Join-Path $directory $file
        $workbook.SaveAs($path, 51)
        try { $workbook.Close($false) }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook); $workbook = $null }
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            producerUiLanguageId = [int]$application.LanguageSettings.LanguageID(2)
            producerDecimalSeparator = [string]$application.International(3)
            producerThousandsSeparator = [string]$application.International(4)
            hostCulture = [Globalization.CultureInfo]::CurrentCulture.Name
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotMultiFieldValueOracle.ps1'
            file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = "Source!A1:C$($rows.Count + 1)"; filteredField = $case.Field; axis = $case.Axis
            filterType = $case.Type; threshold = $case.Threshold; threshold2 = $case.Threshold2
            outputRange = $range; grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-value-multifield-$($case.Key)-conformance.provenance.json") -Encoding utf8
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
