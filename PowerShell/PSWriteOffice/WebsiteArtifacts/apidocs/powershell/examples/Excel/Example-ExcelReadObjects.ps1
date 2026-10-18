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
        Add-OfficeExcelTable -Data $rows -TableName 'Sales' -AutoFit
        Add-OfficeExcelValidationList -Range 'C2:C3' -Values 'New','In Progress','Done'
        Set-OfficeExcelNamedRange -Name 'SalesData' -Range 'A1:C3'
        Set-OfficeExcelFormula -Address 'D2' -Formula 'SUM(A2:C2)'
        Set-OfficeExcelHeaderFooter -HeaderCenter 'Demo' -FooterRight 'Page &P of &N'
        Invoke-OfficeExcelAutoFit -Columns
    }
}

$data = Get-OfficeExcelData -Path $path -Sheet 'Data'
$data | Format-Table

Write-Host 'Named ranges:'
Get-OfficeExcelNamedRange -Path $path -Sheet 'Data' | Format-Table

Write-Host 'Tables:'
Get-OfficeExcelTable -Path $path | Format-Table
