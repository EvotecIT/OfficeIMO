<#
.SYNOPSIS
Creates independent Excel fixtures for three-field pivot value filters.
#>
[CmdletBinding()]
param(
    [ValidateSet('top1', 'top2', 'bottom1', 'greater15', 'between15and30',
        'mixed-row-greater50', 'mixed-row-between45and65', 'mixed-row-top1', 'mixed-row-top1-tie', 'mixed-row-bottom1',
        'mixed-column-bottom1', 'mixed-column-greater85', 'mixed-column-top1',
        'mixed-row-top1-error', 'mixed-column-top1-error',
        'mixed-metric-top1-two-measures', 'mixed-units-top1-two-measures',
        'mixed-metric-row-values-top1-two-measures', 'mixed-units-row-values-top1-two-measures',
        'mixed-metric-column-values-first-top1-two-measures', 'mixed-units-column-values-first-top1-two-measures',
        'mixed-metric-row-values-first-top1-two-measures', 'mixed-units-row-values-first-top1-two-measures',
        'mixed-metric-row-values-middle-top1-two-measures', 'mixed-units-row-values-middle-top1-two-measures',
        'mixed-row-top1-column-bottom1', 'mixed-column-bottom1-row-top1')]
    [string] $Kind = 'top1',
    [string] $OutputDirectory
)
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = if ($OutputDirectory) { $OutputDirectory }
    else { Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus' }
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$directory = (Resolve-Path -LiteralPath $directory).Path

if (-not ('OfficeIMOExcelPivotThreeLevelOracleProcess' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OfficeIMOExcelPivotThreeLevelOracleProcess {
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
    [void][OfficeIMOExcelPivotThreeLevelOracleProcess]::GetWindowThreadProcessId(
        [IntPtr][long]$application.Hwnd, [ref]$excelProcessId)
    if ($excelProcessId -eq 0 -or $existingExcelIds -contains [int]$excelProcessId) {
        throw 'Could not prove isolated Excel instance.'
    }
    $isolated = $true
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $workbook = $application.Workbooks.Add()
    $source = $workbook.Worksheets.Item(1)
    $source.Name = 'Source'
    $twoMeasures = $Kind.EndsWith('-two-measures', [StringComparison]::Ordinal)
    $rowValues = $Kind.Contains('-row-values-', [StringComparison]::Ordinal)
    $valuesFirst = $Kind.Contains('-values-first-', [StringComparison]::Ordinal)
    $valuesMiddle = $Kind.Contains('-values-middle-', [StringComparison]::Ordinal)
    $headers = @('Region', 'Product', 'Channel', 'Sales')
    if ($twoMeasures) { $headers += 'Units' }
    for ($column = 0; $column -lt $headers.Count; $column++) {
        $source.Cells.Item(1, $column + 1).Value2 = $headers[$column]
    }
    $rows = if ($Kind -ne 'top1') {
        @(
            @('East', 'A', 'Retail', 10.0), @('East', 'A', 'Online', 30.0), @('East', 'A', 'Partner', 20.0),
            @('East', 'B', 'Retail', 40.0), @('East', 'B', 'Online', 20.0), @('East', 'B', 'Partner', 5.0),
            @('West', 'A', 'Retail', 5.0), @('West', 'A', 'Online', 25.0), @('West', 'A', 'Partner', 15.0),
            @('West', 'B', 'Retail', 35.0), @('West', 'B', 'Online', 15.0), @('West', 'B', 'Partner', 45.0)
        )
    } else {
        @(
            @('East', 'A', 'Retail', 10.0), @('East', 'A', 'Online', 30.0),
            @('East', 'B', 'Retail', 40.0), @('East', 'B', 'Online', 20.0),
            @('West', 'A', 'Retail', 5.0), @('West', 'A', 'Online', 25.0),
            @('West', 'B', 'Retail', 35.0), @('West', 'B', 'Online', 15.0)
        )
    }
    if ($Kind -eq 'mixed-row-top1-tie') { $rows[0][3] = 15.0 }
    if ($Kind -in @('mixed-row-top1-error', 'mixed-column-top1-error')) { $rows[3][3] = '=1/0' }
    for ($row = 0; $row -lt $rows.Count; $row++) {
        for ($column = 0; $column -lt 4; $column++) {
            if ($column -eq 3) {
                if ($rows[$row][$column] -is [string]) {
                    $source.Cells.Item($row + 2, $column + 1).Formula = [string]$rows[$row][$column]
                } else {
                    $source.Cells.Item($row + 2, $column + 1).Value2 = [double]$rows[$row][$column]
                }
            } else {
                $source.Cells.Item($row + 2, $column + 1).Value2 = [string]$rows[$row][$column]
            }
        }
        if ($twoMeasures) {
            [double]$unitValue = if ($row -in @(3, 4, 5, 6, 7, 8)) { 10.0 } else { 1.0 }
            $source.Cells.Item($row + 2, 5).Value2 = $unitValue
        }
    }
    $view = $workbook.Worksheets.Add()
    $view.Name = 'Grouped'
    $lastSourceRow = $rows.Count + 1
    $lastSourceColumn = if ($twoMeasures) { 5 } else { 4 }
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R$($lastSourceRow)C$lastSourceColumn", 6)
    $pivot = $cache.CreatePivotTable($view.Range('A4'), 'ValuePivot')
    $mixedAxes = $Kind.StartsWith('mixed-', [StringComparison]::Ordinal)
    foreach ($fieldName in @('Region', 'Product', 'Channel')) {
        $field = $pivot.PivotFields($fieldName)
        $field.Orientation = if ($mixedAxes -and $fieldName -eq 'Channel') { 2 } else { 1 }
        $field.Position = if ($fieldName -eq 'Region') { 1 }
            elseif ($fieldName -eq 'Product') { 2 }
            elseif ($mixedAxes) { 1 }
            else { 3 }
    }
    $metric = $pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
    if ($twoMeasures) {
        $units = $pivot.AddDataField($pivot.PivotFields('Units'), 'UnitsMetric', -4157)
        $valuesField = $pivot.DataPivotField
        [int]$valuesOrientation = if ($rowValues) { 1 } else { 2 }
        [int]$valuesPosition = if ($valuesFirst) { 1 } elseif ($valuesMiddle) { 2 } elseif ($rowValues) { 3 } else { 2 }
        $valuesField.Orientation = $valuesOrientation
        $valuesField.Position = $valuesPosition
    }
    $pivot.RowAxisLayout(1)
    [void]$pivot.RefreshTable()
    # Type IDs are XlPivotFilterType values: top=1, bottom=2, greater=9, between=13.
    $rule = switch ($Kind) {
        'top1'          { @{ Type = 1; First = 1.0; Grand = 130.0 } }
        'top2'          { @{ Type = 1; First = 2.0; Grand = 230.0 } }
        'bottom1'       { @{ Type = 2; First = 1.0; Grand = 35.0 } }
        'greater15'     { @{ Type = 9; First = 15.0; Grand = 215.0 } }
        'between15and30' { @{ Type = 13; First = 15.0; Second = 30.0; Grand = 125.0 } }
        'mixed-row-greater50' { @{ Type = 9; First = 50.0; Grand = 220.0; Field = 'Product' } }
        'mixed-row-between45and65' { @{ Type = 13; First = 45.0; Second = 65.0; Grand = 170.0; Field = 'Product' } }
        'mixed-row-top1' { @{ Type = 1; First = 1.0; Grand = 160.0; Field = 'Product' } }
        'mixed-row-top1-tie' { @{ Type = 1; First = 1.0; Grand = 225.0; Field = 'Product' } }
        'mixed-row-bottom1' { @{ Type = 2; First = 1.0; Grand = 105.0; Field = 'Product' } }
        'mixed-column-bottom1' { @{ Type = 2; First = 1.0; Grand = 85.0; Field = 'Channel' } }
        'mixed-column-greater85' { @{ Type = 9; First = 85.0; Grand = 180.0; Field = 'Channel' } }
        'mixed-column-top1' { @{ Type = 1; First = 1.0; Grand = 180.0; Field = 'Channel' } }
        'mixed-row-top1-error' { @{ Type = 1; First = 1.0; Grand = 155.0; Field = 'Product' } }
        'mixed-column-top1-error' { @{ Type = 1; First = 1.0; Grand = 90.0; Field = 'Channel' } }
        'mixed-metric-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 160.0; UnitsGrand = 33.0; Field = 'Product'; Measure = 'Metric' } }
        'mixed-units-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 110.0; UnitsGrand = 60.0; Field = 'Product'; Measure = 'UnitsMetric' } }
        'mixed-metric-row-values-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 160.0; UnitsGrand = 33.0; Field = 'Product'; Measure = 'Metric' } }
        'mixed-units-row-values-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 110.0; UnitsGrand = 60.0; Field = 'Product'; Measure = 'UnitsMetric' } }
        'mixed-metric-column-values-first-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 160.0; UnitsGrand = 33.0; Field = 'Product'; Measure = 'Metric' } }
        'mixed-units-column-values-first-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 110.0; UnitsGrand = 60.0; Field = 'Product'; Measure = 'UnitsMetric' } }
        'mixed-metric-row-values-first-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 160.0; UnitsGrand = 66.0; Field = 'Product'; Measure = 'Metric' } }
        'mixed-units-row-values-first-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 265.0; UnitsGrand = 60.0; Field = 'Product'; Measure = 'UnitsMetric' } }
        'mixed-metric-row-values-middle-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 160.0; UnitsGrand = 66.0; Field = 'Product'; Measure = 'Metric' } }
        'mixed-units-row-values-middle-top1-two-measures' { @{ Type = 1; First = 1.0; Grand = 265.0; UnitsGrand = 60.0; Field = 'Product'; Measure = 'UnitsMetric' } }
        'mixed-row-top1-column-bottom1' { @{ Grand = 35.0 } }
        'mixed-column-bottom1-row-top1' { @{ Grand = 65.0 } }
    }
    $rules = if ($Kind -eq 'mixed-row-top1-column-bottom1') {
        @(
            @{ Type = 1; First = 1.0; Field = 'Product' },
            @{ Type = 2; First = 1.0; Field = 'Channel' }
        )
    } elseif ($Kind -eq 'mixed-column-bottom1-row-top1') {
        @(
            @{ Type = 2; First = 1.0; Field = 'Channel' },
            @{ Type = 1; First = 1.0; Field = 'Product' }
        )
    } else { @($rule) }
    foreach ($savedRule in $rules) {
        $filteredField = if ($savedRule.ContainsKey('Field')) { $savedRule.Field } else { 'Channel' }
        $filterMeasure = if ($savedRule.ContainsKey('Measure') -and $savedRule.Measure -eq 'UnitsMetric') { $units } else { $metric }
        if ($savedRule.ContainsKey('Second')) {
            [void]$pivot.PivotFields($filteredField).PivotFilters.Add2(
                $savedRule.Type, $filterMeasure, $savedRule.First, $savedRule.Second)
        } else {
            [void]$pivot.PivotFields($filteredField).PivotFilters.Add2(
                $savedRule.Type, $filterMeasure, $savedRule.First)
        }
    }
    $lookups = $workbook.Worksheets.Add()
    $lookups.Name = 'Lookups'
    $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
    for ($index = 0; $index -lt $rows.Count; $index++) {
        $lookups.Cells.Item($index + 2, 2).Formula =
            '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","' + $rows[$index][0] +
            '","Product","' + $rows[$index][1] + '","Channel","' + $rows[$index][2] + '")'
    }
    if ($twoMeasures) {
        $lookups.Cells.Item(1, 3).Formula = '=GETPIVOTDATA("UnitsMetric",Grouped!$A$4)'
        for ($index = 0; $index -lt $rows.Count; $index++) {
            $lookups.Cells.Item($index + 2, 3).Formula =
                '=GETPIVOTDATA("UnitsMetric",Grouped!$A$4,"Region","' + $rows[$index][0] +
                '","Product","' + $rows[$index][1] + '","Channel","' + $rows[$index][2] + '")'
        }
    }
    $application.CalculateFullRebuild()
    $range = $pivot.TableRange1.Address($false, $false)
    $grand = [double]$pivot.GetPivotData('Metric').Value2
    if ($grand -ne $rule.Grand) { throw "Excel pivot oracle changed: grand total $grand." }
    if ($twoMeasures) {
        $unitsGrand = [double]$pivot.GetPivotData('UnitsMetric').Value2
        if ($unitsGrand -ne $rule.UnitsGrand) { throw "Excel pivot oracle changed: Units grand total $unitsGrand." }
    }
    $file = "pivot-value-three-level-$Kind-conformance.xlsx"
    $path = Join-Path $directory $file
    $workbook.SaveAs($path, 51)
    try { $workbook.Close($false) }
    finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook); $workbook = $null }
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        producerUiLanguageId = [int]$application.LanguageSettings.LanguageID(2)
        hostCulture = [Globalization.CultureInfo]::CurrentCulture.Name
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotThreeLevelValueOracle.ps1'
        file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = "Source!A1:$([char](64 + $lastSourceColumn))$lastSourceRow"
        measureNames = if ($twoMeasures) { @('Metric', 'UnitsMetric') } else { @('Metric') }
        valuesOnRows = $rowValues
        valuesPosition = if ($twoMeasures) { $valuesPosition } else { $null }
        selectedMeasure = if ($twoMeasures) { $rule.Measure } else { 'Metric' }
        unitsGrandTotal = if ($twoMeasures) { $unitsGrand } else { $null }
        errorCells = if ($Kind -in @('mixed-row-top1-error', 'mixed-column-top1-error')) { 'Source!D5' } else { $null }
        errorFormula = if ($Kind -in @('mixed-row-top1-error', 'mixed-column-top1-error')) { '=1/0' } else { $null }
        filteredField = if (@($rules).Count -eq 1) { $filteredField } else { 'Multiple' }
        axis = if ($mixedAxes) { 'Mixed' } else { 'Row' }
        filterType = if (@($rules).Count -eq 1) { $rule.Type } else { $null }
        threshold = if (@($rules).Count -eq 1) { $rule.First } else { $null }
        threshold2 = if ($rule.ContainsKey('Second')) { $rule.Second } else { $null }
        filters = @($rules | ForEach-Object {
            [ordered]@{ field = if ($_.ContainsKey('Field')) { $_.Field } else { 'Channel' }
                type = $_.Type; threshold = $_.First
                threshold2 = if ($_.ContainsKey('Second')) { $_.Second } else { $null } }
        })
        outputRange = $range; grandTotal = $grand
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-value-three-level-$Kind-conformance.provenance.json") -Encoding utf8
    [pscustomobject]@{ File = $file; Range = $range; Grand = $grand }
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
