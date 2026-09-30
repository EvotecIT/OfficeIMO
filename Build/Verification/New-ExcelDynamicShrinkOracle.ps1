<#
.SYNOPSIS
Creates the independent saved-cache fixture for loaded dynamic-array recalculation.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus'
$path = Join-Path $directory 'dynamic-shrink-source.xlsx'
$mutex = [Threading.Mutex]::new($false, 'Local\OfficeIMO.Excel.Tests.DesktopCom')
$acquired = $false
$application = $null
$workbook = $null
$sheet = $null
$version = $null
$build = $null
try {
    $acquired = $mutex.WaitOne([TimeSpan]::FromMinutes(5))
    if (-not $acquired) { throw 'Excel COM lock timed out.' }
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $version = $application.Version
    $build = $application.Build
    $workbook = $application.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Data'
    $sheet.Range('A1').Value2 = 2
    $sheet.Range('G1').Formula2 = '=SEQUENCE(A1,2)'
    $application.CalculateFullRebuild()
    $values = @($sheet.Range('G1').Text, $sheet.Range('H1').Text,
        $sheet.Range('G2').Text, $sheet.Range('H2').Text)
    if (($values -join ',') -ne '1,2,3,4') {
        throw "Unexpected Excel spill: $($values -join ',')"
    }
    $workbook.SaveAs($path, 51)
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($comObject in @($sheet, $workbook, $application)) {
        if ($null -ne $comObject) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) }
    }
    if ($acquired) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}

[ordered]@{
    producer = 'Microsoft Excel'
    version = $version
    build = $build
    generatedUtc = [DateTime]::UtcNow.ToString('o')
    file = 'dynamic-shrink-source.xlsx'
    sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
    source = [ordered]@{ sheet = 'Data'; input = 'A1=2'; anchor = 'G1'; formula = '=SEQUENCE(A1,2)'; cachedValues = @('1', '2', '3', '4') }
    regeneration = 'Build/Verification/New-ExcelDynamicShrinkOracle.ps1'
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'dynamic-shrink-source.provenance.json') -Encoding utf8
