<#
.SYNOPSIS
Creates independent function/reference/date-system workbooks using desktop Excel.
.DESCRIPTION
Uses a separate hidden Excel instance and closes every owned workbook. Normal
tests consume the checked-in cached results without requiring Office automation.
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$oracleDirectory = Join-Path $repositoryRoot 'OfficeIMO.TestAssets/Documents/ExcelFormulaCorpus'
$formulas = @(
    'SUM(Amounts)', 'SUM(Amounts)+ROUND(1.255,2)', 'AVERAGE(Amounts)',
    'IF(SUM(Amounts)>50,SUMIF(Keys,"Beta",Amounts),0)',
    'COUNTIF(Keys,"*a*")', 'SUMIF(Amounts,">20",Amounts)',
    'SUMIFS(Amounts,Keys,"*a*",Amounts,">=20")', 'COUNTIFS(Keys,"*a*",Amounts,"<40")',
    'AVERAGEIF(Amounts,">=20",Amounts)', 'MINIFS(Amounts,Keys,"*a*")', 'MAXIFS(Amounts,Keys,"*a*")',
    'INDEX(Amounts,MATCH("gamma",Keys,0))', 'VLOOKUP("Beta",''Source Data''!A1:B4,2,FALSE)',
    'XLOOKUP("gamma",Keys,Amounts)', 'XLOOKUP("missing",Keys,Amounts,"fallback")',
    'IFERROR(VLOOKUP("missing",''Source Data''!A1:B4,2,FALSE),99)',
    'ROUND(SUM(Amounts)*TaxRate,2)', 'LocalAmount+Chosen', '''O''''Brien''!A1*2',
    'SUM(''Source Data''!B1:B4)+''O''''Brien''!A1',
    'CONCAT("total=",SUM(Amounts))', 'TEXTJOIN("|",TRUE,Keys)',
    'UPPER(LEFT(INDEX(Keys,2),2))&LOWER(RIGHT(INDEX(Keys,3),2))',
    'SUBSTITUTE(TRIM("  Alpha   Beta  ")," ","_")',
    'LEN(REPT("ab",3))', 'EXACT(LOWER("ALPHA"),INDEX(Keys,1))',
    'FIND("a","Alpha",2)', 'SEARCH("B?ta","Alpha Beta")',
    'VALUE("123.5")+1', 'TEXT(1234.5,"0.00")',
    'ROUND(-1.255,2)', 'ROUNDUP(-1.251,2)', 'ROUNDDOWN(-1.259,2)',
    'TRUNC(-2.9)', 'INT(-2.9)', 'MOD(-7,3)', 'MROUND(7,2)',
    'CEILING.MATH(-2.5)', 'FLOOR.MATH(-2.5)', 'SQRT(81)+POWER(2,3)',
    'SUMPRODUCT(Amounts,Amounts)', 'MEDIAN(Amounts)', 'LARGE(Amounts,2)',
    'PERCENTILE.INC(Amounts,0.25)', 'STDEV.P(Amounts)',
    'ROUND(PMT(0.01,12,1000),2)', 'ROUND(FV(0.01,12,-100),2)',
    'IF(AND(1<2,OR(FALSE,TRUE)),"yes","no")', 'IFS(1>2,"bad",2=2,"ok")',
    'SWITCH(2,1,"one",2,"two","other")', 'CHOOSE(2,"one","two","three")',
    'IFERROR(1/0,"caught")', 'IFNA(NA(),"missing")', 'ISERROR(NA())', 'ISNA(NA())',
    'IFERROR(SUM(''Source Data''!D1),77)', 'COUNT(''Source Data''!D1)', 'COUNTA(''Source Data''!D1)',
    'ISNUMBER(SUM(Amounts))', 'ISLOGICAL(1=1)', 'ISFORMULA(''Source Data''!D1)',
    'DATE(2024,2,29)', 'DATE(2024,2,29)+1', 'YEAR(DATE(2024,2,29))',
    'MONTH(DATE(2024,2,29))', 'DAY(DATE(2024,2,29))',
    'DATE(1900,2,28)', 'DATE(1900,2,29)', 'DATE(1900,3,1)', 'DATE(1904,1,1)',
    'YEAR(60)', 'MONTH(60)', 'DAY(60)',
    'TIME(12,30,15)', 'HOUR(TIME(12,30,15))', 'MINUTE(TIME(12,30,15))', 'SECOND(TIME(12,30,15))',
    'EDATE(DATE(2024,1,31),1)', 'EOMONTH(DATE(2024,1,15),1)',
    'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,5))', 'WORKDAY(DATE(2024,1,5),1)',
    'DATEDIF(DATE(2020,1,1),DATE(2024,1,1),"y")',
    'YEARFRAC(DATE(2024,1,1),DATE(2025,1,1),3)',
    'ISNUMBER(TRUE)', 'ISNUMBER(1=1)', 'ISTEXT(NA())', 'ISLOGICAL("true")',
    'ROUND(2.675,2)', 'ROUND(1.005,2)', 'ROUND(-1.005,2)',
    'WEEKDAY(1)', 'WEEKDAY(59)', 'WEEKDAY(60)', 'WEEKDAY(61)',
    'DAY(0)', 'MONTH(0)', 'YEAR(0)',
    'DATE(1900,1,0)', 'DATE(1900,3,0)', 'DATE(1900,2,30)',
    'SEARCH("B~?ta","B?ta")', 'SEARCH("a*","xxabc")', 'SEARCH("*","Alpha")',
    'SEARCH("~*","x*y")', 'COUNTIF(Keys,"~*")',
    'DATE(2024.9,2.9,29.9)', 'ROUND(1E30,-2)',
    'TaxAlias*100', 'TextConstant&"!"', 'ISLOGICAL(FlagConstant)',
    'IFNA(ErrorConstant,"named missing")', 'Shadowed', '''O''''Brien''!Shadowed',
    'CommaConstant', 'RefTextConstant', 'CommaAlias', 'ISERROR(RefErrorConstant)',
    'DAYS(60,59)', 'DAYS(61,59)', 'DAYS(59.9,60.1)', 'DATEDIF(59,61,"d")',
    'YEARFRAC(59,61,0)', 'YEARFRAC(59,61,1)', 'YEARFRAC(59,61,2)', 'YEARFRAC(59,61,3)', 'YEARFRAC(59,61,4)',
    'NETWORKDAYS(58,62)', 'WORKDAY(59,1)', 'WORKDAY.INTL(59,1,1)',
    'EDATE(60,1)', 'EOMONTH(59,0)', 'DAYS360(59,61)', 'DATEDIF(59,61,"md")',
    'EOMONTH(60,0)', 'EDATE(59,0)', 'EDATE(60,0)', 'EDATE(31,1)',
    'NETWORKDAYS(59,61,60)', 'WORKDAY(59,1,60)', 'DATEDIF(59,61,"yd")',
    'DATEDIF(60,425,"y")', 'DATEDIF(60,425,"m")', 'DATEDIF(60,425,"md")',
    'YEARFRAC(59,60,1)', 'YEARFRAC(60,61,1)',
    'YEARFRAC(DATE(2023,2,28),DATE(2023,3,1),4)',
    'YEARFRAC(DATE(2024,2,28),DATE(2024,3,1),4)',
    'YEARFRAC(DATE(2024,2,29),DATE(2024,3,2),4)',
    'HOUR(0.999999)', 'MINUTE(0.999999)', 'SECOND(0.999999)',
    'DATEDIF(60,60,"yd")', 'DATEDIF(60,59,"d")',
    'WEEKNUM(7)', 'WEEKNUM(60)', 'WEEKNUM(61)', 'WEEKNUM(7,2)',
    'ISOWEEKNUM(7)', 'ISOWEEKNUM(60)', 'ISOWEEKNUM(61)',
    'TEXT(60,"[$-409]mm-dd")', 'TEXT(59,"[$-409]dddd")', 'TEXT(0,"[$-409]mm-dd")', 'YEARFRAC(60,59,1)',
    'DATEDIF(60,425,"yd")', 'NETWORKDAYS(59.9,59.1)'
)

function Set-OracleCell($sheet, [int] $row, [int] $column, $value, [switch] $Formula) {
    $cell = $sheet.Cells.Item($row, $column)
    try {
        if ($Formula) { $cell.Formula = $value }
        elseif ($value -is [double]) { $cell.Value2 = [double]$value }
        else { $cell.Value2 = [string]$value }
    } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell) }
}

$application = $null
$workbooks = $null
$records = @()
try {
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    $application.UseSystemSeparators = $false
    $application.DecimalSeparator = '.'
    $application.ThousandsSeparator = ','
    $workbooks = $application.Workbooks
    foreach ($date1904 in @($false, $true)) {
        $workbook = $null
        $sheets = $null
        $expressions = $null
        $source = $null
        $quoted = $null
        $names = $null
        $localNames = $null
        $quotedNames = $null
        $closed = $false
        try {
            $workbook = $workbooks.Add()
            $workbook.Date1904 = $date1904
            $sheets = $workbook.Worksheets
            $expressions = $sheets.Item(1)
            $expressions.Name = 'Expressions'
            $source = $sheets.Add()
            $source.Name = 'Source Data'
            $quoted = $sheets.Add()
            $quoted.Name = "O'Brien"
            $keys = @('alpha', 'Beta', 'gamma', 'Delta')
            for ($row = 1; $row -le 4; $row++) {
                Set-OracleCell $source $row 1 $keys[$row - 1]
                Set-OracleCell $source $row 2 ([double]($row * 10))
            }
            Set-OracleCell $source 1 4 '=NA()' -Formula
            Set-OracleCell $quoted 1 1 7.0
            $names = $workbook.Names
            foreach ($binding in @(
                @('Keys', "='Source Data'!`$A`$1:`$A`$4"),
                @('Amounts', "='Source Data'!`$B`$1:`$B`$4"),
                @('TaxRate', '=0.2'),
                @('Chosen', "='O''Brien'!`$A`$1"),
                @('TaxAlias', '=TaxRate'), @('TextConstant', '="say ""hello"""'),
                @('FlagConstant', '=TRUE'), @('ErrorConstant', '=#N/A'), @('Shadowed', '=100'),
                @('CommaConstant', '="Hello, world"'), @('RefTextConstant', '="contains #REF!"'),
                @('CommaAlias', '=CommaConstant'), @('RefErrorConstant', '=#REF!'), @('TextExpression', '="a"&"b"')
            )) {
                $name = $names.Add($binding[0], $binding[1])
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($name)
            }
            $localNames = $expressions.Names
            $name = $localNames.Add('LocalAmount', "='Source Data'!`$B`$2")
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($name)
            $name = $localNames.Add('Shadowed', '=5')
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($name)
            $quotedNames = $quoted.Names
            $name = $quotedNames.Add('Shadowed', '=7')
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($name)
            for ($index = 0; $index -lt $formulas.Count; $index++) {
                Set-OracleCell $expressions ($index + 1) 1 $formulas[$index]
                Set-OracleCell $expressions ($index + 1) 2 ('=' + $formulas[$index]) -Formula
            }
            Set-OracleCell $expressions 1 3 '=CONCAT(TextExpression)' -Formula
            $application.CalculateFullRebuild()
            $dateSystem = if ($date1904) { '1904' } else { '1900' }
            $file = 'function-conformance-' + $dateSystem + '.xlsx'
            $path = Join-Path $oracleDirectory $file
            $workbook.SaveAs($path, 51)
            $workbook.Close($false)
            $closed = $true
            $records += [ordered]@{
                file = $file
                dateSystem = $dateSystem
                sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
                cases = $formulas.Count
                sourceFormulaCells = 1
                unsupportedNamedExpressionCells = 1
            }
        } finally {
            if ($null -ne $workbook -and -not $closed) { $workbook.Close($false) }
            foreach ($owned in @($quotedNames, $localNames, $names, $quoted, $source, $expressions, $sheets, $workbook)) {
                if ($null -ne $owned) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned) }
            }
        }
    }
    [ordered]@{
        producer = 'Microsoft Excel'
        version = $application.Version
        build = $application.Build
        generatedUtc = [DateTime]::UtcNow.ToString('o')
        regeneration = 'Build/Verification/New-ExcelFormulaConformanceOracle.ps1'
        numberTextSeparators = 'decimal dot, thousands comma'
        dateTextLocale = 'Explicit [$-409] English weekday names and month/day patterns; localized year tokens are excluded'
        workbooks = $records
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $oracleDirectory 'function-conformance.provenance.json') -Encoding utf8
} finally {
    if ($null -ne $application) { $application.Quit() }
    foreach ($owned in @($workbooks, $application)) {
        if ($null -ne $owned) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned) }
    }
}
