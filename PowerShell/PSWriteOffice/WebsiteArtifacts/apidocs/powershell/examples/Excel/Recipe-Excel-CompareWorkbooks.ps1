$baseline = '.\Excel-Baseline.xlsx'
$candidate = '.\Excel-Candidate.xlsx'

New-OfficeExcel -Path $baseline {
    ExcelSheet 'Data' {
        ExcelCell -Address A1 -Value 'Status'
        ExcelCell -Address A2 -Value 'Draft'
    }
}

New-OfficeExcel -Path $candidate {
    ExcelSheet 'Data' {
        ExcelCell -Address A1 -Value 'Status'
        ExcelCell -Address A2 -Value 'Ready'
    }
}

Compare-OfficeExcelWorkbook -Path $baseline -DifferencePath $candidate
