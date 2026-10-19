$path = '.\Excel-Object-Composition.xlsx'
$projects = @(
    [pscustomobject]@{ Project = 'Atlas'; Owner = 'Operations'; Progress = 0.80 }
    [pscustomobject]@{ Project = 'Beacon'; Owner = 'Security'; Progress = 0.55 }
)

$workbook = New-OfficeExcel -Path $path -NoSave
$sheet = $workbook | Add-OfficeExcelSheet -Name 'Projects' -PassThru
$sheet | Set-OfficeExcelCell -Address A1 -Value 'Delivery portfolio' -BackgroundColor '#D9EAF7'
Add-OfficeExcelTable -Worksheet $sheet -InputObject $projects -StartRow 3 -TableName 'Projects' -AutoFit
Set-OfficeExcelCell -Document $workbook -Sheet 'Projects' -Address C4 -NumberFormat '0%'
$workbook | Save-OfficeExcel
$workbook | Close-OfficeExcel
