Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$data = @(
    [PSCustomObject]@{ Region = 'North America'; Revenue = 125000; YoY = 0.12 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000; YoY = 0.22 }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000; YoY = 0.18 }
)

$path = Join-Path $documents 'Excel-Basic.xlsx'
New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Summary' -Content {
        Set-OfficeExcelRow -Row 1 -Values 'Region', 'Revenue', 'YoY'
        Add-OfficeExcelTable -Data $data -TableName 'Sales' -TableStyle 'TableStyleMedium9'
        Set-OfficeExcelColumn -Column 1 -AutoFit
    }
}

Write-Host "Workbook saved to $path"
