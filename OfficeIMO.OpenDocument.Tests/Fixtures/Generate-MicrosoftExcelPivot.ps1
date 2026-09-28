$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Output "Excel producer version: $($excel.Version)"
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Data'
    $sheet.Range('A1').Value2 = 'Region'
    $sheet.Range('B1').Value2 = 'Month'
    $sheet.Range('C1').Value2 = 'Sales'
    $sheet.Range('A2').Value2 = 'North'
    $sheet.Range('B2').Value2 = 'Jan'
    $sheet.Range('C2').Value2 = 10.0
    $sheet.Range('A3').Value2 = 'North'
    $sheet.Range('B3').Value2 = 'Feb'
    $sheet.Range('C3').Value2 = 20.0
    $sheet.Range('A4').Value2 = 'South'
    $sheet.Range('B4').Value2 = 'Jan'
    $sheet.Range('C4').Value2 = 30.0
    $sheet.Range('A5').Value2 = 'South'
    $sheet.Range('B5').Value2 = 'Feb'
    $sheet.Range('C5').Value2 = 40.0
    $cache = $book.PivotCaches().Create(1, $sheet.Range('A1:C5'))
    $pivot = $cache.CreatePivotTable($sheet.Range('E1'), 'SalesPivot')
    $pivot.PivotFields('Region').Orientation = 1
    $pivot.PivotFields('Month').Orientation = 2
    [void] $pivot.AddDataField($pivot.PivotFields('Sales'), 'Sum of Sales', -4157)
    $xlsxPath = Join-Path $PSScriptRoot 'microsoft-excel-pivot.xlsx'
    $book.SaveAs($xlsxPath, 51)
    Write-Output $xlsxPath
    $outputPath = Join-Path $PSScriptRoot 'microsoft-excel-pivot.ods'
    $book.SaveAs($outputPath, 60)
    Write-Output $outputPath
} finally {
    if ($book -ne $null) { $book.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
