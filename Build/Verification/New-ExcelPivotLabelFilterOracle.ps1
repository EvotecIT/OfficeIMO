<#
.SYNOPSIS
Creates independent Microsoft Excel label-filter pivot fixtures.
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
        [pscustomobject]@{ Key = 'equal'; Type = 15; Criterion = 'East'; Total = 60.0 },
        [pscustomobject]@{ Key = 'not-equal'; Type = 16; Criterion = 'East'; Total = 155.0 },
        [pscustomobject]@{ Key = 'begins'; Type = 17; Criterion = 'North'; Total = 70.0 },
        [pscustomobject]@{ Key = 'not-begins'; Type = 18; Criterion = 'North'; Total = 145.0 },
        [pscustomobject]@{ Key = 'ends'; Type = 19; Criterion = 'east'; Total = 115.0 },
        [pscustomobject]@{ Key = 'not-ends'; Type = 20; Criterion = 'east'; Total = 100.0 },
        [pscustomobject]@{ Key = 'not-contains'; Type = 22; Criterion = 'east'; Total = 100.0 }
    )
    foreach ($case in $cases) {
        $workbook = $application.Workbooks.Add()
        $source = $workbook.Worksheets.Item(1)
        $source.Name = 'Source'
        $rows = @(
            @('Region', 'Sales'), @('East', 60.0), @('North', 30.0),
            @('Northeast', 40.0), @('Southeast', 15.0), @('West', 70.0)
        )
        $values = New-Object 'object[,]' 6, 2
        for ($row = 0; $row -lt 6; $row++) {
            for ($column = 0; $column -lt 2; $column++) { $values[$row, $column] = $rows[$row][$column] }
        }
        $source.Range('A1:B6').Value2 = $values
        $view = $workbook.Worksheets.Add()
        $view.Name = 'Grouped'
        $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R6C2", 6)
        $pivot = $cache.CreatePivotTable($view.Range('A4'), 'LabelPivot')
        $region = $pivot.PivotFields('Region')
        $region.Orientation = 1
        $region.Position = 1
        [void]$pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
        $pivot.RowAxisLayout(1)
        [void]$pivot.RefreshTable()
        [void]$region.PivotFilters.Add2($case.Type, [Type]::Missing, $case.Criterion)
        $lookups = $workbook.Worksheets.Add()
        $lookups.Name = 'Lookups'
        $names = @('East', 'North', 'Northeast', 'Southeast', 'West')
        $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
        for ($index = 0; $index -lt $names.Count; $index++) {
            $lookups.Cells.Item($index + 2, 2).Formula =
                '=GETPIVOTDATA("Metric",Grouped!$A$4,"Region","' + $names[$index] + '")'
        }
        $application.CalculateFullRebuild()
        $range = $pivot.TableRange1.Address($false, $false)
        $grand = $pivot.GetPivotData('Metric').Value2
        if ($grand -ne $case.Total) { throw "Unexpected $($case.Key) grand total: $grand" }
        $file = "pivot-label-$($case.Key)-conformance.xlsx"
        $path = Join-Path $directory $file
        $workbook.SaveAs($path, 51)
        $workbook.Close($false)
        $workbook = $null
        [ordered]@{
            producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
            generatedUtc = [DateTime]::UtcNow.ToString('o')
            regeneration = 'Build/Verification/New-ExcelPivotLabelFilterOracle.ps1'
            file = $file
            sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            sourceRange = 'Source!A1:B6'; filterType = $case.Type; criterion = $case.Criterion
            outputRange = $range; grandTotal = $grand
        } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory "pivot-label-$($case.Key)-conformance.provenance.json") -Encoding utf8
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
