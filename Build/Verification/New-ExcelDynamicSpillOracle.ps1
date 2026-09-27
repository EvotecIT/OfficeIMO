<#
.SYNOPSIS
Creates independent dynamic-spill lifecycle cases using a separate desktop Excel instance.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$directory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus'
$path = Join-Path $directory 'dynamic-spills.xlsx'
$application = $null
$workbooks = $null
$workbook = $null
$closed = $false
$observations = [Collections.Generic.List[object]]::new()

function Invoke-OracleRange {
    param($Sheet, [string] $Reference, [scriptblock] $Action)
    $range = $Sheet.Range($Reference)
    try { & $Action $range } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
}

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
    $names = @('Initial', 'Grown', 'Shrunk', 'Blocked', 'Recovered', 'Merged', 'Table', 'Edge', 'EmptyFilter', 'EmptyUnique', 'ZeroRows', 'NegativeRows', 'ZeroColumns', 'NegativeColumns')
    for ($i = 0; $i -lt $names.Count; $i++) {
        $sheet = if ($i -eq 0) { $workbook.Worksheets.Item(1) } else { $workbook.Worksheets.Add() }
        try {
            $sheet.Name = $names[$i]
            Invoke-OracleRange $sheet 'A1' { param($range) $range.Value2 = 2.0 }
            Invoke-OracleRange $sheet 'B1' { param($range) $range.Value2 = 2.0 }
            Invoke-OracleRange $sheet 'G1' { param($range) $range.Formula2 = '=SEQUENCE(A1,B1)' }
            switch ($sheet.Name) {
                'Grown' {
                    $application.CalculateFullRebuild()
                    Invoke-OracleRange $sheet 'A1' { param($range) $range.Value2 = 4.0 }
                }
                'Shrunk' {
                    Invoke-OracleRange $sheet 'A1' { param($range) $range.Value2 = 4.0 }
                    $application.CalculateFullRebuild()
                    Invoke-OracleRange $sheet 'A1' { param($range) $range.Value2 = 1.0 }
                }
                'Blocked' {
                    Invoke-OracleRange $sheet 'G2' { param($range) $range.Formula = "'blocker" }
                    Invoke-OracleRange $sheet 'I1' { param($range) $range.Formula = '=G1' }
                }
                'Recovered' {
                    Invoke-OracleRange $sheet 'G2' { param($range) $range.Formula = "'blocker" }
                    $application.CalculateFullRebuild()
                    Invoke-OracleRange $sheet 'G2' { param($range) [void]$range.ClearContents() }
                }
                'Merged' {
                    Invoke-OracleRange $sheet 'G2:H2' { param($range) [void]$range.Merge() }
                }
                'Table' {
                    Invoke-OracleRange $sheet 'G2' { param($range) $range.Formula = "'First" }
                    Invoke-OracleRange $sheet 'H2' { param($range) $range.Formula = "'Second" }
                    $range = $sheet.Range('G2:H3')
                    $tables = $sheet.ListObjects
                    $table = $null
                    try { $table = $tables.Add(1, $range, $null, 1); $table.Name = 'SpillBlocker' }
                    finally {
                        foreach ($comObject in @($table, $tables, $range)) {
                            if ($null -ne $comObject) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) }
                        }
                    }
                }
                'Edge' {
                    Invoke-OracleRange $sheet 'G1' { param($range) [void]$range.ClearContents() }
                    Invoke-OracleRange $sheet 'XFD1048576' { param($range) $range.Formula2 = '=SEQUENCE(2,2)' }
                }
                'EmptyFilter' {
                    Invoke-OracleRange $sheet 'A1:A4' { param($range) $range.Value2 = 1.0 }
                    Invoke-OracleRange $sheet 'D1:D4' { param($range) $range.Value2 = 0.0 }
                    Invoke-OracleRange $sheet 'G1' { param($range) $range.Formula2 = '=FILTER(A1:A4,D1:D4)' }
                }
                'EmptyUnique' {
                    Invoke-OracleRange $sheet 'A1:A4' { param($range) $range.Value2 = 1.0 }
                    Invoke-OracleRange $sheet 'G1' { param($range) $range.Formula2 = '=UNIQUE(A1:A4,FALSE,TRUE)' }
                    Invoke-OracleRange $sheet 'I1' { param($range) $range.Formula = '=G1' }
                }
                'ZeroRows' { Invoke-OracleRange $sheet 'G1' { param($range) $range.Formula2 = '=SEQUENCE(0)' } }
                'NegativeRows' { Invoke-OracleRange $sheet 'G1' { param($range) $range.Formula2 = '=SEQUENCE(-1)' } }
                'ZeroColumns' { Invoke-OracleRange $sheet 'G1' { param($range) $range.Formula2 = '=SEQUENCE(1,0)' } }
                'NegativeColumns' { Invoke-OracleRange $sheet 'G1' { param($range) $range.Formula2 = '=SEQUENCE(1,-1)' } }
            }
            $application.CalculateFullRebuild()
            $anchor = if ($sheet.Name -eq 'Edge') { 'XFD1048576' } else { 'G1' }
            $range = $sheet.Range($anchor)
            try {
                $observations.Add([ordered]@{ sheet = $sheet.Name; anchor = $anchor; formula = $range.Formula2; text = $range.Text })
            } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
        } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
    }
    $workbook.SaveAs($path, 51)
    $workbook.Close($false)
    $closed = $true
    [ordered]@{
        producer = 'Microsoft Excel'
        version = $application.Version
        build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        file = 'dynamic-spills.xlsx'
        sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        observations = $observations
        regeneration = 'Build/Verification/New-ExcelDynamicSpillOracle.ps1'
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'dynamic-spills.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
    if ($null -ne $application) { $application.Quit() }
    foreach ($comObject in @($workbook, $workbooks, $application)) {
        if ($null -ne $comObject) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) }
    }
}
