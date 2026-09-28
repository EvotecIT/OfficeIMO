<#
.SYNOPSIS
Creates Excel-produced SORT collation cases in a separate desktop Excel instance.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$outputDirectory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus'
$existingExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excel = $null
$workbook = $null
$isolated = $false
$closed = $false
$caseCount = 0
$workbookPath = Join-Path $outputDirectory 'sort-collation.xlsx'
$jsonPath = Join-Path $outputDirectory 'sort-collation.provenance.json'

function Add-OracleCase($name, $values, $formula, $outputRows, $outputColumns) {
    $sheet = if ($script:caseCount -eq 0) { $script:workbook.Worksheets.Item(1) } else { $script:workbook.Worksheets.Add() }
    $script:caseCount++
    $sheet.Name = $name
    for ($row = 0; $row -lt $values.Count; $row++) {
        for ($column = 0; $column -lt $values[$row].Count; $column++) {
            if ($null -ne $values[$row][$column]) {
                $address = '{0}{1}' -f [char]([int][char]'A' + $column), ($row + 1)
                $cellValue = $values[$row][$column]
                $cell = $sheet.Range($address)
                if ($cellValue -is [int] -or $cellValue -is [double]) {
                    $cell.Formula = '={0}' -f $cellValue
                } elseif ($cellValue -is [bool]) {
                    $cell.Formula = if ($cellValue) { '=TRUE()' } else { '=FALSE()' }
                } else {
                    $cell.Value2 = [string]$cellValue
                }
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
            }
        }
    }
    $anchor = $sheet.Range('G1')
    $anchor.Formula2 = $formula
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($anchor)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
    return [ordered]@{ name=$name; formula=$formula; rows=$outputRows; columns=$outputColumns }
}

try {
    $excel = New-Object -ComObject Excel.Application
    Start-Sleep -Milliseconds 500
    $newExcelIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $existingExcelIds -notcontains $_.Id } |
        ForEach-Object { $_.Id })
    if ($newExcelIds.Count -ne 1) {
        throw "Excel COM did not produce one independently identified process; new process count: $($newExcelIds.Count)."
    }
    $isolated = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) {
        $extra = $workbook.Worksheets.Item($workbook.Worksheets.Count)
        try { $extra.Delete() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($extra) }
    }
    $cases = @(
        (Add-OracleCase 'TextAscending' @(@('pear',1),@('apple',2),@('banana',3)) '=SORT(A1:B3,1,1)' 3 2),
        (Add-OracleCase 'TextDescending' @(@('pear',1),@('apple',2),@('banana',3)) '=SORT(A1:B3,1,-1)' 3 2),
        (Add-OracleCase 'BlankNumeric' @(@(2,1),@($null,2),@(1,3)) '=SORT(A1:B3,1,1)' 3 2),
        (Add-OracleCase 'BlankNumericDescending' @(@(2,1),@($null,2),@(1,3)) '=SORT(A1:B3,1,-1)' 3 2),
        (Add-OracleCase 'MixedKeys' @(@(2,1),@('b',2),@(1,3),@('a',4)) '=SORT(A1:B4,1,1)' 4 2),
        (Add-OracleCase 'MixedDescending' @(@(2,1),@('b',2),@(1,3),@('a',4)) '=SORT(A1:B4,1,-1)' 4 2),
        (Add-OracleCase 'MixedBoolean' @(@($false,1),@('a',2),@(1,3),@($true,4),@($null,5)) '=SORT(A1:B5,1,1)' 5 2),
        (Add-OracleCase 'MixedBooleanDescending' @(@($false,1),@('a',2),@(1,3),@($true,4),@($null,5)) '=SORT(A1:B5,1,-1)' 5 2),
        (Add-OracleCase 'BlankText' @(@('pear',1),@($null,2),@('apple',3)) '=SORT(A1:B3,1,1)' 3 2),
        (Add-OracleCase 'BlankTextDescending' @(@('pear',1),@($null,2),@('apple',3)) '=SORT(A1:B3,1,-1)' 3 2),
        (Add-OracleCase 'TextByColumns' @(@('pear','apple','banana'),@(1,2,3)) '=SORT(A1:C2,1,1,TRUE)' 2 3),
        (Add-OracleCase 'MixedByColumnsDescending' @(@($false,'a',1,$true,$null),@(1,2,3,4,5)) '=SORT(A1:E2,1,-1,TRUE)' 2 5),
        (Add-OracleCase 'TextCaseTies' @(@('apple',1),@('Apple',2),@('APPLE',3)) '=SORT(A1:B3,1,1)' 3 2)
    )
    $excel.CalculateFullRebuild()
    $results = @()
    foreach ($case in $cases) {
        $sheet = $workbook.Worksheets.Item($case.name)
        $output = @()
        for ($row = 1; $row -le $case.rows; $row++) {
            $values = @()
            for ($column = 1; $column -le $case.columns; $column++) {
                $address = '{0}{1}' -f [char]([int][char]'G' + $column - 1), $row
                $cell = $sheet.Range($address)
                try { $values += $cell.Value2 }
                finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
            }
            $output += ,$values
        }
        $results += [ordered]@{name=$case.name;formula=$case.formula;output=$output}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
    }
    $workbook.SaveAs($workbookPath, 51)
    $workbook.Close($false)
    $closed = $true
    $evidence = [ordered]@{
        producer='Microsoft Excel'
        version=$excel.Version
        build=$excel.Build
        generatedUtc=[DateTime]::UtcNow.ToString('o')
        file='sort-collation.xlsx'
        sha256=(Get-FileHash -LiteralPath $workbookPath -Algorithm SHA256).Hash.ToLowerInvariant()
        inputNote='Numeric and Boolean source cells are Excel-calculated constants; text cells are literals.'
        cases=$results
        regeneration='Build/Verification/New-ExcelSortOracle.ps1'
    }
    [IO.File]::WriteAllText($jsonPath,($evidence | ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
    Write-Output "Created $($evidence.file) with $($cases.Count) Excel-calculated cases; SHA-256 $($evidence.sha256)."
}
finally {
    try {
        if ($workbook -ne $null) {
            try {
                if (-not $closed) { $workbook.Close($false) }
            } finally {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            }
        }
    } finally {
        if ($excel -ne $null) {
            try {
                if ($isolated) { $excel.Quit() }
            } finally {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            }
        }
    }
}
