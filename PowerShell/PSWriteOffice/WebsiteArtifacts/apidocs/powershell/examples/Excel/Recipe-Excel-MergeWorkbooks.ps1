$target = '.\Excel-Consolidated.xlsx'
$source = '.\Excel-Regional-Source.xlsx'

New-OfficeExcel -Path $target {
    ExcelSheet 'Summary' {
        ExcelCell -Address A1 -Value 'Consolidated service report'
    }
}

New-OfficeExcel -Path $source {
    ExcelSheet 'North' {
        ExcelCell -Address A1 -Value 'North region'
        ExcelCell -Address B1 -Value 42
    }
    ExcelSheet 'South' {
        ExcelCell -Address A1 -Value 'South region'
        ExcelCell -Address B1 -Value 37
    }
}

Copy-OfficeExcelSheet `
    -Path $target `
    -Source $source `
    -Sheet 'North', 'South' `
    -NamePrefix 'Region '
