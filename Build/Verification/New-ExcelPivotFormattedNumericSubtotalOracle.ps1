<#
.SYNOPSIS
Creates a Microsoft Excel pivot fixture with formatted numeric subtotals.
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
    $workbook = $application.Workbooks.Add()
    $source = $workbook.Worksheets.Item(1)
    $source.Name = 'Source'
    $headers = @('Quantity', 'Product', 'Sales')
    for ($column = 0; $column -lt $headers.Count; $column++) {
        $source.Cells.Item(1, $column + 1).Value2 = $headers[$column]
    }
    $rows = @(
        @(1000, 'A', 10), @(1000, 'B', 20), @(2000, 'A', 30)
    )
    for ($row = 0; $row -lt $rows.Count; $row++) {
        $source.Cells.Item($row + 2, 1).Value2 = [int]$rows[$row][0]
        $source.Cells.Item($row + 2, 2).Value2 = [string]$rows[$row][1]
        $source.Cells.Item($row + 2, 3).Value2 = [int]$rows[$row][2]
    }
    $source.Range('A2:A4').NumberFormat = '#,##0'
    $view = $workbook.Worksheets.Add()
    $view.Name = 'Grouped'
    $cache = $workbook.PivotCaches().Create(1, "'Source'!R1C1:R4C3", 6)
    $pivot = $cache.CreatePivotTable($view.Range('A4'), 'LabelPivot')
    $outer = $pivot.PivotFields('Quantity')
    $outer.Orientation = 1
    $outer.Position = 1
    $inner = $pivot.PivotFields('Product')
    $inner.Orientation = 1
    $inner.Position = 2
    [void]$pivot.AddDataField($pivot.PivotFields('Sales'), 'Metric', -4157)
    $pivot.RowAxisLayout(1)
    $pivot.SubtotalLocation(2)
    [void]$pivot.RefreshTable()
    $lookups = $workbook.Worksheets.Add()
    $lookups.Name = 'Lookups'
    $lookups.Cells.Item(1, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4)'
    $lookups.Cells.Item(2, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4,"Quantity",1000)'
    $lookups.Cells.Item(3, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4,"Quantity",2000)'
    $lookups.Cells.Item(4, 2).Formula = '=GETPIVOTDATA("Metric",Grouped!$A$4,"Quantity",1000,"Product","A")'
    $application.CalculateFullRebuild()
    $range = $pivot.TableRange1.Address($false, $false)
    $firstSubtotal = [string]$view.Cells.Item(7, 1).Text
    $secondSubtotal = [string]$view.Cells.Item(9, 1).Text
    $grand = [double]$pivot.GetPivotData('Metric').Value2
    if ($range -ne 'A4:C10' -or $firstSubtotal -ne '1,000 Total' -or
        $secondSubtotal -ne '2,000 Total' -or $grand -ne 60.0) {
        throw "Unexpected Excel subtotal view: range=$range first=$firstSubtotal second=$secondSubtotal grand=$grand"
    }
    $file = 'pivot-label-number-subtotal-conformance.xlsx'
    $path = Join-Path $directory $file
    $workbook.SaveAs($path, 51)
    $outer.RepeatLabels = $true
    [void]$pivot.RefreshTable()
    $application.CalculateFullRebuild()
    $repeatLabel = [string]$view.Cells.Item(6, 1).Text
    if ($repeatLabel -ne '1,000') { throw "Unexpected repeated Excel label: $repeatLabel" }
    $repeatFile = 'pivot-label-number-subtotal-repeat-conformance.xlsx'
    $repeatPath = Join-Path $directory $repeatFile
    $workbook.SaveAs($repeatPath, 51)
    $workbook.Close($false)
    $workbook = $null
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotFormattedNumericSubtotalOracle.ps1'
        file = $file; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:C4'; numberFormat = '#,##0'
        outputRange = $range; firstSubtotal = $firstSubtotal; secondSubtotal = $secondSubtotal; grandTotal = $grand
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'pivot-label-number-subtotal-conformance.provenance.json') -Encoding utf8
    [pscustomobject]@{ File = $file; Range = $range; FirstSubtotal = $firstSubtotal; SecondSubtotal = $secondSubtotal; Grand = $grand }
    [ordered]@{
        producer = 'Microsoft Excel'; version = $application.Version; build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelPivotFormattedNumericSubtotalOracle.ps1'
        file = $repeatFile; sha256 = (Get-FileHash -LiteralPath $repeatPath -Algorithm SHA256).Hash.ToLowerInvariant()
        sourceRange = 'Source!A1:C4'; numberFormat = '#,##0'; repeatLabels = $true
        outputRange = $range; repeatedLabel = $repeatLabel; grandTotal = $grand
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'pivot-label-number-subtotal-repeat-conformance.provenance.json') -Encoding utf8
    [pscustomobject]@{ File = $repeatFile; Range = $range; RepeatedLabel = $repeatLabel; Grand = $grand }
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $application) {
        $application.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
