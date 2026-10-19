Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Excel-Pictures.xlsx'
$imagePath = Join-Path $PSScriptRoot '..\Word\Example-WordTableCells.fixture.png'

New-OfficeExcel -Path $path {
    ExcelSheet 'Pictures' {
        ExcelCell -Address 'A1' -Value 'Range anchored image'
        ExcelCell -Address 'E1' -Value 'Scaled image'
        ExcelCell -Address 'E8' -Value 'Rotated image'

        ExcelImage -Path $imagePath -Range 'A2:C12' -Name 'HeaderLogo' -AltText 'Company logo pinned to A2 through C12' -Placement MoveAndSize
        ExcelImage -Path $imagePath -Address 'E2' -ScalePercent 20 -Name 'ScaledLogo' -AltText 'Logo scaled to 20 percent'
        ExcelImage -Path $imagePath -Address 'E9' -Width 120 -Height 48 -RotationDegrees 12 -Name 'RotatedLogo' -AltText 'Logo rotated by 12 degrees'

        ExcelColumn -ColumnName A -Width 18
        ExcelColumn -ColumnName B -Width 18
        ExcelColumn -ColumnName C -Width 18
        ExcelColumn -ColumnName E -Width 24
    }
}

Write-Host "Workbook saved to $path"
