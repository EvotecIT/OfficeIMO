$csv = '.\Regional-Sales.csv'
$workbook = '.\Regional-Sales.xlsx'

ExcelNew -Path $workbook {
    ExcelSheet 'Readme' {
        ExcelCell -Address A1 -Value 'Imported from Regional-Sales.csv'
    }
}

Import-OfficeExcelDelimitedText `
    -Path $workbook `
    -SourcePath $csv `
    -Delimiter ';' `
    -SheetName 'Sales'
