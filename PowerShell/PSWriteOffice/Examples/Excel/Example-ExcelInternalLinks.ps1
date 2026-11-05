Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Excel-InternalLinks.xlsx'
$rows = @(
    [PSCustomObject]@{ Sheet = 'Alpha'; Target = 'Alpha' }
    [PSCustomObject]@{ Sheet = 'Beta'; Target = 'Beta' }
)

New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Summary' -Content {
        Add-OfficeExcelTable -InputObject $rows -Name 'SummaryTable'
        Set-OfficeExcelCell -Address 'D1' -Value 'Sheet'
        Set-OfficeExcelCell -Address 'D2' -Value 'Alpha'
        Set-OfficeExcelCell -Address 'D3' -Value 'Beta'

        Set-OfficeExcelHyperlink -Internal -Range 'D2:D3'
        Set-OfficeExcelHyperlink -Internal -Header 'Sheet' -Table 'SummaryTable' -DisplayScript { param($text) "Open $text" }
    }
    Add-OfficeExcelSheet -Name 'Alpha' -Content {
        Set-OfficeExcelCell -Address 'A1' -Value 'Alpha Home'
    }
    Add-OfficeExcelSheet -Name 'Beta' -Content {
        Set-OfficeExcelCell -Address 'A1' -Value 'Beta Home'
    }
}

Write-Host "Workbook saved to $path"
