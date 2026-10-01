<#
.SYNOPSIS
Creates the bounded array calculation oracle with a separate desktop Excel instance.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus'
$path = Join-Path $directory 'bounded-arrays.xlsx'
$cases = @(
    'SEQUENCE(3,2,10,-2)', 'SEQUENCE(2.9,2,0,0)', 'SEQUENCE(0)',
    'FILTER(A1:B4,D1:D4)', 'FILTER(A1:B4,E1:E4,"empty")',
    'FILTER(A1:B4,E1:E4)', 'FILTER(A1:B4,A1:A4>1)',
    'SORT(A1:B4,1,1)', 'SORT(A1:B4,1,-1)', 'SORT(A1:B4,3)',
    'UNIQUE(A1:B4)', 'UNIQUE(A1:B4,FALSE,TRUE)', 'UNIQUE(A1:B4,TRUE)',
    'SORT(UNIQUE(A1:B4))', 'FILTER(A1:B4,F1:F4)',
    'UNIQUE(A1:A4)', 'FILTER(A1:B4,D6:E6)', 'SORT(A1:B4,1,-1,TRUE)',
    'FILTER(A1:A2,A1:A2>0,"empty")', 'SORT(A1:B4,1.9)'
)
$application = $null
$workbooks = $null
$workbook = $null
$closed = $false
try {
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $workbooks = $application.Workbooks
    $workbook = $workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) {
        $extra = $workbook.Worksheets.Item($workbook.Worksheets.Count)
        try { $extra.Delete() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($extra) }
    }
    for ($i = 0; $i -lt $cases.Count; $i++) {
        $sheet = if ($i -eq 0) { $workbook.Worksheets.Item(1) } else { $workbook.Worksheets.Add() }
        try {
            $sheet.Name = 'Case' + ($i + 1)
            $numbers = @(3.0, 1.0, 3.0, 2.0)
            $texts = @('c', 'a', 'c', 'b')
            for ($row = 1; $row -le 4; $row++) {
                foreach ($column in @(1, 2, 4, 5, 6)) {
                    $cell = $sheet.Cells.Item($row, $column)
                    try {
                        switch ($column) {
                            1 { $cell.Value2 = $numbers[$row - 1] }
                            2 { $cell.Formula = "'" + $texts[$row - 1] }
                            4 { $cell.Formula = if ($row -ne 2) { '=TRUE()' } else { '=FALSE()' } }
                            5 { $cell.Value2 = 0.0 }
                            6 { $cell.Value2 = 1.0 }
                        }
                    } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                }
            }
            if ($i -eq 15) {
                foreach ($row in @(1, 2, 3, 4)) {
                    $cell = $sheet.Cells.Item($row, 1)
                    try {
                        switch ($row) {
                            1 { $cell.Value2 = 1.0 }
                            2 { $cell.Formula = "'1" }
                            3 { $cell.Formula = '=TRUE()' }
                            4 { $cell.ClearContents() | Out-Null }
                        }
                    } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                }
            }
            if ($i -eq 16) {
                foreach ($column in @(4, 5)) {
                    $cell = $sheet.Cells.Item(6, $column)
                    try { $cell.Formula = if ($column -eq 4) { '=TRUE()' } else { '=FALSE()' } }
                    finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                }
            }
            if ($i -eq 17) {
                for ($row = 1; $row -le 4; $row++) {
                    $cell = $sheet.Cells.Item($row, 2)
                    try { $cell.Value2 = [double](9 + $row) }
                    finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                }
            }
            if ($i -eq 14) {
                $cell = $sheet.Cells.Item(2, 6)
                try { $cell.Formula = '=NA()' } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
            }
            if ($i -eq 18) {
                foreach ($row in @(1, 2)) {
                    $cell = $sheet.Cells.Item($row, 1)
                    try { $cell.Value2 = if ($row -eq 1) { 5e-8 } else { -5e-8 } }
                    finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                }
                foreach ($column in @(8, 9, 10)) {
                    $cell = $sheet.Cells.Item(1, $column)
                    try { $cell.Formula = switch ($column) { 8 { '=COUNTIF(A1:A2,0)' } 9 { '=MATCH(0,A1:A2,0)' } 10 { '=RANK.AVG(A1,A1:A2)' } } }
                    finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
                }
            }
            $cell = $sheet.Cells.Item(1, 7)
            try { $cell.Formula2 = '=' + $cases[$i] } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
        } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
    }
    $application.CalculateFullRebuild()
    $workbook.SaveAs($path, 51)
    $workbook.Close($false)
    $closed = $true
    [ordered]@{
        producer = 'Microsoft Excel'
        version = $application.Version
        build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        file = 'bounded-arrays.xlsx'
        sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        formulas = $cases
        regeneration = 'Build/Verification/New-ExcelArrayFormulaOracle.ps1'
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'bounded-arrays.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($comObject in @($workbook, $workbooks, $application)) {
        if ($null -ne $comObject) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) }
    }
}
