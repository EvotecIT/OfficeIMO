Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Excel-LinksAndImages.xlsx'

New-OfficeExcel -Path $path {
    Add-OfficeExcelSheet -Name 'Data' -Content {
        Set-OfficeExcelCell -Address 'A1' -Value 'Reference'
        Set-OfficeExcelCell -Address 'B1' -Value 'Host'

        Set-OfficeExcelSmartHyperlink -Address 'A2' -Url 'https://datatracker.ietf.org/doc/html/rfc7208'
        Set-OfficeExcelHostHyperlink -Address 'B2' -Url 'https://learn.microsoft.com/office/open-xml/'
        Add-OfficeExcelImageFromUrl -Address 'D2' -Url 'https://raw.githubusercontent.com/github/explore/main/topics/powershell/powershell.png' -WidthPixels 48 -HeightPixels 48
    }
}

Write-Host "Workbook saved to $path"
