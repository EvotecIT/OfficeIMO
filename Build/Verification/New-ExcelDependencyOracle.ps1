<#
.SYNOPSIS
Creates independently calculated dependency and circular-reference fixtures.
.DESCRIPTION
Uses a separate hidden Excel instance with iteration disabled. Circular caches
are producer artifacts, not an expected numeric solution to an unsupported cycle.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus'
$application = $null
$workbooks = $null
$records = @()
try {
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $workbooks = $application.Workbooks
    foreach ($kind in @('dependency', 'circular')) {
        $workbook = $null
        $sheets = $null
        $sheet = $null
        $closed = $false
        try {
            $workbook = $workbooks.Add()
            $application.Iteration = $false
            $sheets = $workbook.Worksheets
            $sheet = $sheets.Item(1)
            $sheet.Name = if ($kind -eq 'dependency') { 'Chains' } else { 'Circular' }
            if ($kind -eq 'dependency') {
                # A runs in worksheet order; B requires the reverse dependency order.
                $values = New-Object 'object[,]' 301, 2
                $values[0, 0] = 1.0
                $values[300, 1] = 1.0
                for ($row = 1; $row -le 300; $row++) {
                    $values[$row, 0] = '=A' + $row + '+1'
                    $values[($row - 1), 1] = '=B' + ($row + 1) + '+1'
                }
                $range = $sheet.Range('A1:B301')
                try { $range.Formula = $values }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
            } else {
                $values = New-Object 'object[,]' 1, 4
                $values[0, 0] = '=B1+1'
                $values[0, 1] = '=A1+1'
                $values[0, 2] = '=SUM(A1:B1)'
                $values[0, 3] = '=40+2'
                $range = $sheet.Range('A1:D1')
                try { $range.Formula = $values }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
            }
            $application.CalculateFullRebuild()
            $file = $kind + '-conformance.xlsx'
            $path = Join-Path $directory $file
            $workbook.SaveAs($path, 51)
            $workbook.Close($false)
            $closed = $true
            $records += [ordered]@{
                file = $file
                sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
                formulaCells = if ($kind -eq 'dependency') { 600 } else { 4 }
                maximumAcyclicDepth = if ($kind -eq 'dependency') { 300 } else { 1 }
            }
        } finally {
            if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
            foreach ($owned in @($sheet, $sheets, $workbook)) {
                if ($null -ne $owned) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned) }
            }
        }
    }
    [ordered]@{
        producer = 'Microsoft Excel'
        version = $application.Version
        build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelDependencyOracle.ps1'
        iterationEnabled = $false
        circularContract = 'Report cycle and its dependent; clear guarded caches and retain formulas. Producer cycle caches are not numeric conformance expectations.'
        workbooks = $records
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'dependency-conformance.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $application) { $application.Quit() }
    foreach ($owned in @($workbooks, $application)) {
        if ($null -ne $owned) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned) }
    }
}
