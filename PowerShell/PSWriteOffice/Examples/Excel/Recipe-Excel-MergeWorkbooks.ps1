$target = '.\Excel-Consolidated.xlsx'
$source = '.\Excel-Regional-Source.xlsx'

ExcelNew -Path $target {
    ExcelSheet 'Summary' {
        ExcelCell -Address A1 -Value 'Consolidated service report'
    }
}

ExcelNew -Path $source {
    ExcelSheet 'North' {
        ExcelCell -Address A1 -Value 'North region'
        ExcelCell -Address B1 -Value 42
    }
    ExcelSheet 'South' {
        ExcelCell -Address A1 -Value 'South region'
        ExcelCell -Address B1 -Value 37
    }
}

Join-OfficeExcelWorkbook `
    -Path $target `
    -SourcePath $source `
    -SourceSheet 'North', 'South' `
    -SheetNamePrefix 'Region '
