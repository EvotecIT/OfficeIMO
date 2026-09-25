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
    $sheet.Cells.Item(1, 1).Value2 = 'Month'
    $sheet.Cells.Item(1, 2).Value2 = 'Sales'
    $sheet.Cells.Item(2, 1).Value2 = 'Jan'
    $sheet.Cells.Item(2, 2).Value2 = 10
    $sheet.Cells.Item(3, 1).Value2 = 'Feb'
    $sheet.Cells.Item(3, 2).Value2 = 20
    $chartObject = $sheet.ChartObjects().Add(180, 40, 400, 250)
    $chartObject.Chart.SetSourceData($sheet.Range('A1:B3'))
    $chartObject.Chart.HasTitle = $true
    $chartObject.Chart.ChartTitle.Text = 'Sales'
    foreach ($variant in @(
        @{ Kind = 'column'; ChartType = 51 },
        @{ Kind = 'bar'; ChartType = 57 },
        @{ Kind = 'line'; ChartType = 4 }
    )) {
        $chartObject.Chart.ChartType = $variant.ChartType
        $outputPath = Join-Path $PSScriptRoot ("microsoft-excel-{0}-chart.ods" -f $variant.Kind)
        $book.SaveAs($outputPath, 60)
        Write-Output $outputPath
    }
} finally {
    if ($book -ne $null) { $book.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
