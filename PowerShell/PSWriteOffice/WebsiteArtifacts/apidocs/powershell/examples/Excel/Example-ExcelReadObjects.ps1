Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Excel-ReadObjects.xlsx'
$rows = @(
    [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
)

New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Data' -Content {
        Add-OfficeExcelTable -InputObject $rows -Name 'Sales'
        Add-OfficeExcelValidation -Type List -Range 'C2:C3' -Values 'New','In Progress','Done'
        Set-OfficeExcelNamedRange -Name 'SalesData' -Range 'A1:C3'
        Set-OfficeExcelCell -Address 'D2' -Formula 'SUM(A2:C2)'
        Set-OfficeExcelHeader -Center 'Demo'
        ExcelFooter -Right 'Page &P of &N'
        Invoke-OfficeExcelAutoFit -Column '*'
    }
}

$data = Import-OfficeExcel -Path $path -Sheet 'Data'
$data | Format-Table

Write-Host 'Named ranges:'
Get-OfficeExcelNamedRange -Path $path -Sheet 'Data' | Format-Table

Write-Host 'Tables:'
Get-OfficeExcelTable -Path $path | Format-Table
