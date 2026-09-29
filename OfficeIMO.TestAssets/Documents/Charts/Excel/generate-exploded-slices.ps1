param([Parameter(Mandatory)][string]$OutputPath)
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Exploded'
    $sheet.Cells.Item(1, 1).Value2 = 'Status'
    $sheet.Cells.Item(1, 2).Value2 = 'Value'
    $sheet.Cells.Item(2, 1).Value2 = 'Complete'
    $sheet.Cells.Item(2, 2).Value2 = 7
    $sheet.Cells.Item(3, 1).Value2 = 'Pending'
    $sheet.Cells.Item(3, 2).Value2 = 3
    $source = $sheet.Range('A1:B3')
    foreach ($spec in @(
        @{ Type = 5; Left = 20; Top = 30; Name = 'Pie exploded'; FileStem = 'exploded-pie-reference' },
        @{ Type = -4120; Left = 470; Top = 30; Name = 'Doughnut exploded'; FileStem = 'exploded-doughnut-reference' }
    )) {
        $chart = $sheet.ChartObjects().Add($spec.Left, $spec.Top, 430, 330).Chart
        $chart.ChartType = $spec.Type
        $chart.SetSourceData($source)
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = $spec.Name
        $series = $chart.SeriesCollection(1)
        $series.Points(1).Explosion = 25
        $series.Points(2).Explosion = 0
        $reference = Join-Path (Split-Path -Parent $OutputPath) ($spec.FileStem + '.png')
        if (-not $chart.Export($reference, 'PNG')) { throw "Excel could not export $reference" }
    }
    $workbook.SaveAs($OutputPath, 51)
} finally {
    if ($workbook -ne $null) { $workbook.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }
    if ($workbook -ne $null) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
    if ($excel -ne $null) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
