$csv = '.\Regional-Sales.csv'
$workbook = '.\Regional-Sales.xlsx'

New-OfficeExcel -Path $workbook {
    ExcelSheet 'Readme' {
        ExcelCell -Address A1 -Value 'Imported from Regional-Sales.csv'
    }
}

Import-OfficeCsv -Path $csv -Delimiter ';' |
    Export-OfficeExcel -Path $workbook -Sheet Sales -Append
