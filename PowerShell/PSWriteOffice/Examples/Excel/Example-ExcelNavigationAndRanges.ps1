Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Excel-NavigationAndRanges.xlsx'
$rows = @(
    [PSCustomObject]@{ Region = 'North America'; Revenue = 125000 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000 }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000 }
)

New-OfficeExcel -Path $path {
    ExcelSheet 'Data' {
        ExcelTable -Data $rows -TableName 'Sales' -AutoFit
        ExcelNamedRange -Name 'SalesData' -Range 'A1:B4'
    }
    ExcelSheet 'Notes' {
        ExcelRow -Row 1 -Values 'Label', 'Value'
        ExcelRow -Row 2 -Values 'Generated', (Get-Date -Format 'yyyy-MM-dd')
    }
    ExcelTableOfContents -IncludeNamedRanges
}

Write-Host "Workbook saved to $path"
Write-Host ''
Write-Host 'Explicit range read:'
Get-OfficeExcelRange -Path $path -Sheet 'Data' -Range 'A1:B4' | Format-Table

Write-Host ''
Write-Host 'Used range read as DataTable:'
$usedRange = Get-OfficeExcelUsedRange -Path $path -Sheet 'Notes' -AsDataTable
$usedRange | Format-Table

Write-Host ''
Write-Host 'TOC rows:'
Get-OfficeExcelRange -Path $path -Sheet 'TOC' -Range 'A3:C5' -AsHashtable | Format-Table
