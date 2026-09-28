<#
.SYNOPSIS
Creates Excel-produced all-error pivot ranking cases and saved-cache provenance.
#>
[CmdletBinding()]
param([string] $OutputDirectory)
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$targetDirectory = if ($OutputDirectory) { $OutputDirectory }
    else { Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelPivotCorpus/AllErrorRanking' }
New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
$targetDirectory = (Resolve-Path -LiteralPath $targetDirectory).Path
$mutex = [Threading.Mutex]::new($false, 'Local\OfficeIMO.Excel.Tests.DesktopCom')
$acquired = $false
$existingExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$excel = $null
$workbook = $null
$isolated = $false
$cases = @(
    @{ Name = 'top-count'; Type = 1; Value = 1.0 },
    @{ Name = 'bottom-count'; Type = 2; Value = 1.0 },
    @{ Name = 'top-count-two'; Type = 1; Value = 2.0 },
    @{ Name = 'bottom-count-two'; Type = 2; Value = 2.0 },
    @{ Name = 'top-count-reversed'; Type = 1; Value = 1.0; ReverseSource = $true },
    @{ Name = 'bottom-count-reversed'; Type = 2; Value = 1.0; ReverseSource = $true },
    @{ Name = 'top-count-same'; Type = 1; Value = 1.0; SameError = $true },
    @{ Name = 'bottom-count-same'; Type = 2; Value = 1.0; SameError = $true },
    @{ Name = 'top-average'; Type = 1; Value = 1.0; Function = -4106 },
    @{ Name = 'bottom-average'; Type = 2; Value = 1.0; Function = -4106 },
    @{ Name = 'top-grouped'; Type = 1; Value = 1.0; GroupedErrors = $true },
    @{ Name = 'bottom-grouped'; Type = 2; Value = 1.0; GroupedErrors = $true },
    @{ Name = 'top-grouped-reversed'; Type = 1; Value = 1.0; GroupedErrors = $true; ReverseSource = $true },
    @{ Name = 'bottom-grouped-reversed'; Type = 2; Value = 1.0; GroupedErrors = $true; ReverseSource = $true }
)
for ($count = 1; $count -le 7; $count++) {
    $cases += @{ Name = "top-full-$count"; Type = 1; Value = [double]$count; FullErrors = $true }
    $cases += @{ Name = "bottom-full-$count"; Type = 2; Value = [double]$count; FullErrors = $true }
}
$results = @()
try {
    $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5))
    if (-not $acquired) { throw 'Excel COM lock timed out.' }
    $excel = New-Object -ComObject Excel.Application
    Start-Sleep -Milliseconds 500
    $newExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $existingExcelIds -notcontains $_.Id } | ForEach-Object Id)
    if ($newExcelIds.Count -ne 1) { throw 'Could not prove isolated Excel instance.' }
    $isolated = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    foreach ($case in $cases) {
        $record = [ordered]@{ name = $case.Name; filterType = $case.Type; value = $case.Value }
        try {
            $workbook = $excel.Workbooks.Add()
            $source = $workbook.Worksheets.Item(1)
            $source.Name = 'Source'
            $source.Range('A1').Value2 = 'Region'
            $source.Range('B1').Value2 = 'Sales'
            $items = @(@('Alpha','=NA()'),@('Bravo','=1/0'),@('Charlie','=VALUE("bad")'))
            if ($case.GroupedErrors) {
                $items = @(
                    @('Alpha','=NA()'),@('Alpha','=1/0'),
                    @('Bravo','=VALUE("bad")'),@('Bravo','=INDIRECT("XFE1")'),
                    @('Charlie','=SQRT(-1)'),@('Charlie','=SUM(A1:A2 B1:B2)')
                )
            }
            if ($case.FullErrors) {
                $items = @(
                    @('Alpha','=NA()'),@('Bravo','=1/0'),@('Charlie','=VALUE("bad")'),
                    @('Delta','=SQRT(-1)'),@('Echo','=INDIRECT("XFE1")'),
                    @('Foxtrot','=BOGUS()'),@('Golf','=SUM(A1:A2 B1:B2)')
                )
            }
            if ($case.ReverseSource) { [array]::Reverse($items) }
            for ($index = 0; $index -lt $items.Count; $index++) {
                $item = $items[$index]
                $row = $index + 2
                $source.Range("A$row").Value2 = $item[0]
                $source.Range("B$row").Formula = if ($case.SameError) { '=NA()' } else { $item[1] }
            }
            $excel.CalculateFullRebuild()
            $view = $workbook.Worksheets.Add()
            $view.Name = 'Grouped'
            $sourceEndRow = $items.Count + 1
            $record.sourceRange = "A1:B$sourceEndRow"
            $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R$($sourceEndRow)C2", 6)
            $pivot = $cache.CreatePivotTable($view.Range('A4'), 'AllErrorPivot')
            $field = $pivot.PivotFields('Region')
            $field.Orientation = 1
            $field.Position = 1
            $function = if ($case.Function) { $case.Function } else { -4157 }
            $record.aggregateFunction = $function
            $measure = $pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', $function)
            $pivot.RowAxisLayout(1)
            [void]$pivot.RefreshTable()
            $filter = $field.PivotFilters.Add2($case.Type, $measure, $case.Value)
            $excel.CalculateFullRebuild()
            $record.outputRange = $pivot.TableRange1.Address($false, $false)
            $viewRows = [int]$pivot.TableRange1.Rows.Count
            $rows = @()
            for ($row = 4; $row -lt 4 + $viewRows; $row++) {
                $rows += [ordered]@{ label = [string]$view.Cells.Item($row,1).Text; metric = [string]$view.Cells.Item($row,2).Text }
            }
            $record.rows = $rows
            $path = Join-Path $targetDirectory ($case.Name + '.xlsx')
            $workbook.SaveAs($path, 51)
            $record.file = [IO.Path]::GetFileName($path)
        } catch {
            $record.error = $_.Exception.Message
        } finally {
            if ($workbook -ne $null) {
                try { $workbook.Close($false) }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook); $workbook = $null }
            }
        }
        if ($record.error) { throw "Excel case '$($case.Name)' failed: $($record.error)" }
        $record.sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        $results += $record
    }
    $manifest = [ordered]@{
        producer = 'Microsoft Excel'
        version = $excel.Version
        build = $excel.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotAllErrorRankingOracle.ps1'
        cases = $results
    }
    [IO.File]::WriteAllText((Join-Path $targetDirectory 'provenance.json'),($manifest | ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
    Write-Output "Created $($results.Count) Excel all-error pivot cases in $targetDirectory."
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
