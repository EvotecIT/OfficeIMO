<#
.SYNOPSIS
Creates independent Microsoft Excel pivot label comparison and range fixtures.
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
        [pscustomobject]@{ Key = 'greater'; Type = 23; From = 'Charlie'; To = $null; Total = 150.0 },
        [pscustomobject]@{ Key = 'greater-equal'; Type = 24; From = 'Charlie'; To = $null; Total = 180.0 },
        [pscustomobject]@{ Key = 'less'; Type = 25; From = 'Charlie'; To = $null; Total = 30.0 },
        [pscustomobject]@{ Key = 'less-equal'; Type = 26; From = 'Charlie'; To = $null; Total = 60.0 },
        [pscustomobject]@{ Key = 'between'; Type = 27; From = 'Bravo'; To = 'Echo'; Total = 140.0 },
        [pscustomobject]@{ Key = 'not-between'; Type = 28; From = 'Bravo'; To = 'Echo'; Total = 70.0 }
    )
    foreach ($case in $cases) {
        $workbook = $application.Workbooks.Add()
        $source = $workbook.Worksheets.Item(1)
        $source.Name = 'Source'
        $rows = @(
            @('Region', 'Sales'), @('Alpha', 10.0), @('Bravo', 20.0),
            @('charlie', 30.0), @('Delta', 40.0), @('Echo', 50.0), @('Foxtrot', 60.0)
        )
        $values = New-Object 'object[,]' 7, 2
        for ($row = 0; $row -lt 7; $row++) {
            for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
        }
        $source.Range('A1:B7').Value2 = $values
        $view = $workbook.Worksheets.Add()
        $view.Name = 'Grouped'
        $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R7C2", 6)
        $pivot = $cache.CreatePivotTable($view.Range('A4'), 'LabelPivot')
        $field = $pivot.PivotFields('Region')
        $field.Orientation = 1
        $field.Position = 1
        [void]$pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
        $pivot.RowAxisLayout(1)
        [void]$pivot.RefreshTable()
        if ($null -eq $case.To) {
            [void]$field.PivotFilters.Add2($case.Type, [Type]::Missing, $case.From)
        } else {
            [void]$field.PivotFilters.Add2($case.Type, [Type]::Missing, $case.From, $case.To)
        }
        $lookups = $workbook.Worksheets.Add()
        $lookups.Name = 'Lookups'
        $names = @('Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo', 'Foxtrot')
        $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
        for ($index = 0; $index -lt $names.Count; $index++) {
            $lookups.Cells.Item($index + 2, 2).Formula =
                '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","' + $names[$index] + '")'
        }
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = $pivot.GetPivotData('Metric').Value2
        if ($grand -ne $case.Total) { throw "Unexpected $($case.Key) grand total: $grand" }
        $file = "pivot-label-range-$($case.Key)-conformance.xlsx"
        $path = Join-Path $directory $file
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
        $workbook = $null
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotLabelRangeOracle.ps1'
            file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = 'Source!A1:B7'; filterType = $case.Type
            value1 = $case.From; value2 = $case.To; outputRange = $range; grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-label-range-$($case.Key)-conformance.provenance.json") -Encoding utf8
        [pscustomobject]@{ File = $file; Range = $range; Grand = $grand }
    }
    # A separate fixture captures Excel's localized ordering where ordinal
    # case-insensitive comparison selects a different item set.
    $workbook = $application.Workbooks.Add()
    $source = $workbook.Worksheets.Item(1)
    $source.Name = 'Source'
    $labels = @('A', 'A-1', 'A1', 'A10', 'A2', 'A_B', 'Ab', 'ä', 'B')
    $source.Cells.Item(1, 1).Value2 = 'Label'
    $source.Cells.Item(1, 2).Value2 = 'Amount'
    for ($index = 0; $index -lt $labels.Count; $index++) {
        $source.Cells.Item($index + 2, 1).Value2 = $labels[$index]
        $source.Cells.Item($index + 2, 2).Value2 = [double][Math]::Pow(2, $index)
    }
    $view = $workbook.Worksheets.Add()
    $view.Name = 'Grouped'
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R10C2", 6)
    $pivot = $cache.CreatePivotTable($view.Range('A4'), 'LabelPivot')
    $field = $pivot.PivotFields('Label')
    $field.Orientation = 1
    [void]$pivot.AddDataField($pivot.PivotFields('Amount'), 'Metric', -4157)
    $pivot.RowAxisLayout(1)
    [void]$pivot.RefreshTable()
    [void]$field.PivotFilters.Add2(23, [Type]::Missing, 'A1')
    $grand = $pivot.GetPivotData('Metric').Value2
    if ($grand -ne 346.0) { throw "Unexpected collation grand total: $grand" }
    $range = $pivot.TableRange1.Address($false, $false)
    $file = 'pivot-label-range-collation-conformance.xlsx'
    $path = Join-Path $directory $file
    $workbook.SaveAs($path, 51)
    $workbook.Close($false)
    $workbook = $null
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotLabelRangeOracle.ps1'
        file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:B10'; filterType = 23; value1 = 'A1'
        labels = $labels; outputRange = $range; grandTotal = $grand; ordinalIgnoreCaseTotal = 504
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'pivot-label-range-collation-conformance.provenance.json') -Encoding utf8
    [pscustomobject]@{ File = $file; Range = $range; Grand = $grand }
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $application) {
        $application.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
