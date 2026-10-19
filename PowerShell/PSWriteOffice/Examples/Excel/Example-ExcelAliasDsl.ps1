Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$orders = @(
    [PSCustomObject]@{ Item = 'Router'; Qty = 15; Status = 'In Stock' }
    [PSCustomObject]@{ Item = 'Switch'; Qty = 4; Status = 'Low' }
    [PSCustomObject]@{ Item = 'Firewall'; Qty = 7; Status = 'In Stock' }
)

$path = Join-Path $documents 'Excel-AliasDsl.xlsx'

New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Inventory' -Content {
        ExcelCell -Address 'A1' -Value 'Item'
        ExcelCell -Address 'B1' -Value 'Qty'
        ExcelCell -Address 'C1' -Value 'Status'

        ExcelTable -Data $orders -TableName 'InventoryTable'
    }
}

Write-Host "Workbook saved to $path"
