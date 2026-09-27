<#
.SYNOPSIS
Creates Excel-authored pivot aggregation and GETPIVOTDATA evidence.
.DESCRIPTION
Each aggregation uses the same worksheet records, row/column keys, tabular layout
and grand totals. All caches and lookup results are produced by native Excel.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus'
[void](New-Item -ItemType Directory -Path $directory -Force)
$functions = [ordered]@{
    Sum = -4157; Count = -4112; CountNumbers = -4113; Average = -4106
    Minimum = -4139; Maximum = -4136; Product = -4149
    StandardDeviation = -4155; StandardDeviationP = -4156
    Variance = -4164; VarianceP = -4165
}
$lookupEdges = @(
    'GETPIVOTDATA("Metric",Sum!A1,"Region","West","Product","A")',
    'GETPIVOTDATA("Amount",Sum!A1)',
    'GETPIVOTDATA("Count",ErrorGroups!A1)', 'GETPIVOTDATA("CountNumbers",ErrorGroups!A1)',
    'GETPIVOTDATA("Metric",Sum!B4,"Region","North")',
    'GETPIVOTDATA("Metric",Sum!A1:D6,"Product","A")',
    'GETPIVOTDATA("Metric",Sum!Z1)',
    'GETPIVOTDATA("Metric",Sum!A1,"Region","North","Region","North")',
    'GETPIVOTDATA("Metric",Sum!A1,"Region","North","Region","South")',
    'GETPIVOTDATA("Metric",Sum!A1,"Region","north")',
    'GETPIVOTDATA("Metric",Sum!A1,"notafield","A")', 'GETPIVOTDATA("missing",Sum!A1)',
    'GETPIVOTDATA("Count",ErrorGroups!A1,"Group","TextNumber")',
    'IFERROR(GETPIVOTDATA("Metric",Sum!Z1),77)',
    'GETPIVOTDATA("Metric",Sum!A1,"Region",Source!A2)',
    'GETPIVOTDATA("Metric",Sum!A1,"Region","West")',
    'GETPIVOTDATA("Metric",Sum!A1,"Region","West","Product","B")',
    'GETPIVOTDATA("Metric",Sum!A1,"Product","a")',
    'GETPIVOTDATA("Count",ErrorGroups!A1,"Group","Literal")',
    'GETPIVOTDATA("CountNumbers",ErrorGroups!A1,"Group","Literal")',
    'GETPIVOTDATA("Metric",Sum!A1,"Product",1)',
    'FALSE()'
)
$application = $null; $workbooks = $null; $workbook = $null
$sheets = $null; $source = $null; $oracles = $null; $caches = $null
$closed = $false
try {
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $workbooks = $application.Workbooks
    $workbook = $workbooks.Add()
    $sheets = $workbook.Worksheets
    $source = $sheets.Item(1)
    $source.Name = 'Source'
    $rows = @(
        @('Region', 'Product', 'Amount', 'Units'),
        @('North', 'A', 10.0, 2.0), @('North', 'B', 20.0, 3.0),
        @('South', 'A', 5.0, 1.0), @('South', 'B', 15.0, 4.0),
        @('North', 'A', -2.0, 1.0), @('West', 'B', $null, 5.0),
        @('South', 'A', 'pending', 2.0), @('North', 'B', $true, 1.0)
    )
    $values = New-Object 'object[,]' $rows.Count, 4
    for ($row = 0; $row -lt $rows.Count; $row++) {
        for ($column = 0; $column -lt 4; $column++) { $values[$row, $column] = $rows[$row][$column] }
    }
    $range = $source.Range('A1:D9')
    try { $range.Value2 = $values }
    finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
    $oracles = $sheets.Add()
    $oracles.Name = 'Lookups'
    $caches = $workbook.PivotCaches()
    $lookupRow = 1
    $records = @()
    foreach ($function in $functions.GetEnumerator()) {
        $sheet = $null; $cache = $null; $pivot = $null
        $rowField = $null; $columnField = $null; $amount = $null; $measure = $null
        try {
            $sheet = $sheets.Add()
            $sheet.Name = $function.Key
            $cache = $caches.Create(1, "'Source'!R1C1:R9C4", 6)
            $destination = $sheet.Range('A1')
            try { $pivot = $cache.CreatePivotTable($destination, ('Pivot' + $function.Key)) }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($destination) }
            $rowField = $pivot.PivotFields('Region')
            $rowField.Orientation = 1
            $rowField.Position = 1
            $rowField.Subtotals = @($false, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false)
            $columnField = $pivot.PivotFields('Product')
            $columnField.Orientation = 2
            $columnField.Position = 1
            $amount = $pivot.PivotFields('Amount')
            $measure = $pivot.AddDataField($amount, 'Metric', $function.Value)
            $pivot.RowAxisLayout(1)
            $pivot.RowGrand = $true
            $pivot.ColumnGrand = $true
            [void]$pivot.RefreshTable()
            foreach ($keys in @(@(), @('Region', 'North'), @('Region', 'South'), @('Region', 'West'),
                @('Product', 'A'), @('Product', 'B'), @('Region', 'North', 'Product', 'A'),
                @('Region', 'South', 'Product', 'B'), @('Region', 'missing'))) {
                $formula = '=GETPIVOTDATA("Metric",' + $function.Key + '!A1'
                foreach ($key in $keys) { $formula += ',"' + $key + '"' }
                $formula += ')'
                $cell = $oracles.Cells.Item($lookupRow, 1)
                try { $cell.Value2 = $function.Key }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                $cell = $oracles.Cells.Item($lookupRow, 2)
                try { $cell.Formula = $formula }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                $lookupRow++
            }
            $tableRange = $pivot.TableRange1
            try { $outputRange = $tableRange.Address($false, $false) }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($tableRange) }
            $records += [ordered]@{ name = 'Pivot' + $function.Key; sheet = $function.Key; function = $function.Key; outputRange = $outputRange }
        } finally {
            foreach ($owned in @($measure, $amount, $columnField, $rowField, $pivot, $cache, $sheet)) {
                if ($null -ne $owned) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned) }
            }
        }
    }
    $errorSource = $null; $errorSheet = $null; $errorCache = $null; $errorPivot = $null; $groupField = $null; $errorAmount = $null
    try {
        $errorSource = $sheets.Add()
        $errorSource.Name = 'ErrorSource'
        $errorValues = New-Object 'object[,]' 8, 2
        $errorRows = @(@('Group', 'Amount'), @('Error', '=NA()'), @('Error', 3.0),
            @('Literal', '="#N/A"'), @('Empty', $null), @('Single', 5.0), @('Boolean', '=TRUE()'), @('TextNumber', '="12"'))
        for ($row = 0; $row -lt 8; $row++) {
            $errorValues[$row, 0] = $errorRows[$row][0]
            $errorValues[$row, 1] = $errorRows[$row][1]
        }
        $range = $errorSource.Range('A1:B8')
        try { $range.Value2 = $errorValues }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
        foreach ($formulaRow in @(2, 4, 7, 8)) {
            $cell = $errorSource.Cells.Item($formulaRow, 2)
            try { $cell.Formula = [string]$errorRows[$formulaRow - 1][1] }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
        }
        $application.CalculateFullRebuild()
        $errorSheet = $sheets.Add()
        $errorSheet.Name = 'ErrorGroups'
        $errorCache = $caches.Create(1, "'ErrorSource'!R1C1:R8C2", 6)
        $destination = $errorSheet.Range('A1')
        try { $errorPivot = $errorCache.CreatePivotTable($destination, 'PivotErrorGroups') }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($destination) }
        $groupField = $errorPivot.PivotFields('Group')
        $groupField.Orientation = 1
        $errorAmount = $errorPivot.PivotFields('Amount')
        foreach ($function in $functions.GetEnumerator()) {
            $measure = $errorPivot.AddDataField($errorAmount, $function.Key, $function.Value)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($measure)
            foreach ($group in @('Error', 'Literal', 'Empty', 'Single', 'Boolean', 'TextNumber')) {
                $cell = $oracles.Cells.Item($lookupRow, 1)
                try { $cell.Value2 = 'Error ' + $function.Key }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                $cell = $oracles.Cells.Item($lookupRow, 2)
                try { $cell.Formula = '=GETPIVOTDATA("' + $function.Key + '",ErrorGroups!A1,"Group","' + $group + '")' }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                $lookupRow++
            }
        }
        [void]$errorPivot.RefreshTable()
    } finally {
        foreach ($owned in @($errorAmount, $groupField, $errorPivot, $errorCache, $errorSheet, $errorSource)) {
            if ($null -ne $owned) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned) }
        }
    }
    $edgeSheet = $null
    try {
        $edgeSheet = $sheets.Add()
        $edgeSheet.Name = 'LookupEdges'
        for ($index = 0; $index -lt $lookupEdges.Count; $index++) {
            $cell = $edgeSheet.Cells.Item($index + 1, 1)
            try { $cell.Formula = '=' + $lookupEdges[$index] }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
        }
    } finally {
        if ($null -ne $edgeSheet) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($edgeSheet) }
    }
    $application.CalculateFullRebuild()
    $path = Join-Path $directory 'aggregation-conformance.xlsx'
    $workbook.SaveAs($path, 51)
    $workbook.Close($false)
    $closed = $true
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotOracle.ps1'
        file = 'aggregation-conformance.xlsx'
        sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:D9'; errorSourceRange = 'ErrorSource!A1:B8'; lookupFormulaCells = $lookupRow - 1
        aggregationLookupCells = 99; typedErrorLookupCells = 66
        lookupEdgeFormulaCells = $lookupEdges.Count
        layout = 'Tabular, one row key, one column key, both grand totals'
        pivots = $records
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'aggregation-conformance.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($owned in @($caches, $oracles, $source, $sheets, $workbook, $workbooks, $application)) {
        if ($null -ne $owned) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned) }
    }
}
