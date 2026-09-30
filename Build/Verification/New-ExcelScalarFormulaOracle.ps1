<#
.SYNOPSIS
Creates the independent scalar-expression test workbook using desktop Excel.
.DESCRIPTION
Requires Windows and installed Excel. Creates and closes a separate hidden Excel
instance. The checked-in workbook lets normal tests use Excel's cached results
without installing or running Office. Run from any working directory.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$oracleDirectory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus'
$oraclePath = Join-Path $oracleDirectory 'scalar-expressions.xlsx'
$formulas = @(
    '1+2*3', '(1+2)*3', '2^3^2', '-2^2', '2^-2', '50%*8',
    'SUM(1,2)*3', 'SUM(1+2,3*4)', 'IF(1+2*3=7,2^3,0)',
    '"a"&"b"', 'LEN("a"&"bc")+1', '1e-3*1000', '1/(2-2)',
    'IFERROR(1/(2-2),99)', '"abc"+1', '#DIV/0!+1', '10^400',
    '(A1+SUM(A1,2))*2', '"a"&1+2', '2^50%', '(-2)^2',
    'SUM(1,1/(2-2))', 'AVERAGE(1,1/(2-2))', 'MIN(1,1/(2-2))',
    'MAX(1,1/(2-2))', 'PRODUCT(1,1/(2-2))', 'SUMSQ(1,1/(2-2))',
    '"x"&1.00', '"x"&1e-3', '"x"&TRUE', '"x"&(1=1)', '"x"&AND(TRUE,TRUE)',
    'COUNT(1,1/(2-2))', '"x"&ISNUMBER(1)', '"x"&FALSE',
    '"x"&EXACT("a","a")', '"x"&EXACT("a","b")', '"x"&ISFORMULA(A1)', '1=1'
)
$application = $null
$workbooks = $null
$workbook = $null
$sheet = $null
$workbookClosed = $false
try {
    if (-not (Test-Path -LiteralPath $oracleDirectory)) {
        New-Item -ItemType Directory -Path $oracleDirectory -ErrorAction Stop | Out-Null
    }
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $application.UseSystemSeparators = $false
    $application.DecimalSeparator = '.'
    $application.ThousandsSeparator = ','
    $workbooks = $application.Workbooks
    $workbook = $workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Expressions'
    $sheet.Cells.Item(1, 1).Value2 = 3.0
    for ($index = 0; $index -lt $formulas.Count; $index++) {
        $cell = $sheet.Cells.Item($index + 1, 2)
        try { $cell.Formula = '=' + $formulas[$index] }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
    }
    $application.CalculateFullRebuild()
    $workbook.SaveAs($oraclePath, 51)
    $workbook.Close($false)
    $workbookClosed = $true
    [ordered]@{
        producer = 'Microsoft Excel'
        version = $application.Version
        build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        file = 'scalar-expressions.xlsx'
        sha256 = (Get-FileHash -LiteralPath $oraclePath -Algorithm SHA256).Hash.ToLowerInvariant()
        cases = $formulas.Count
        regeneration = 'Build/Verification/New-ExcelScalarFormulaOracle.ps1'
        numberTextSeparators = 'decimal dot, thousands comma'
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $oracleDirectory 'provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and -not $workbookClosed) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($comObject in @($sheet, $workbook, $workbooks, $application)) {
        if ($null -ne $comObject) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) }
    }
}
