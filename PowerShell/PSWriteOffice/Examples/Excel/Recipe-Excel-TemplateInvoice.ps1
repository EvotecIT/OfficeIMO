$template = '.\Invoice-Template.xlsx'
$invoice = '.\Invoice-1042.xlsx'

ExcelNew -Path $template {
    ExcelSheet 'Invoice' {
        ExcelCell -Address A1 -Value 'Invoice {{Number}}'
        ExcelCell -Address A3 -Value 'Customer: {{Customer}}'
        ExcelCell -Address A5 -Value 'Amount: {{Amount:currency}}'
        ExcelCell -Address A7 -Value 'Due: {{Due:date}}'
    }
}

Copy-Item -Path $template -Destination $invoice
Invoke-OfficeExcelTemplate -Path $invoice -Value @{
    Number = '1042'
    Customer = 'Northwind Traders'
    Amount = 1840.50
    Due = [datetime]'2026-09-15'
}
