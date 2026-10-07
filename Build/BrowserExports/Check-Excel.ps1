param(
    [Parameter(Mandatory)][string] $WorkbookPath,
    [Parameter(Mandatory)][string] $EvidenceDirectory
)
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $true)
    $sheet = $workbook.Worksheets.Item(1)
    if ($sheet.Cells.Item(4, 1).Value2 -cne '_x0041_') { throw 'Excel changed literal OOXML escape text.' }
    if ($sheet.Cells.Item(3, 1).Value2 -cne '=literal') { throw 'Excel changed formula-like text.' }
    if ($sheet.Cells.Item(2, 4).Value2 -ne 12.5 -or $sheet.Cells.Item(2, 5).Value2 -ne $true) { throw 'Excel changed numeric/boolean values.' }
    if ($sheet.Cells.Item(3, 3).Value2 -ne 59.5 -or $sheet.Cells.Item(4, 3).Value2 -ne 61 -or $sheet.Cells.Item(5, 3).Value2 -ne 1) { throw 'Excel changed 1900 date serials.' }
    if (-not $sheet.Cells.Item(1, 1).Font.Bold -or $sheet.Cells.Item(3, 1).HasFormula) { throw 'Excel header/formula types differ.' }
    if (-not $workbook.Windows.Item(1).FreezePanes) { throw 'Excel did not restore the frozen header.' }
    $originalWidth = $sheet.Columns.Item(1).ColumnWidth
    # The fixture leaves some widths unspecified. Fit this read-only preview so
    # the PDF shows formatted values rather than Excel's narrow-column hashes.
    $sheet.UsedRange.Columns.AutoFit() | Out-Null
    $sheet.UsedRange.Rows.AutoFit() | Out-Null
    $dateText = $sheet.Cells.Item(2, 3).Text
    if ($dateText -ne '2026-10-05 12:34') { throw 'Excel formatted date differs.' }
    $sheet.ExportAsFixedFormat(0, (Join-Path (Resolve-Path -LiteralPath $EvidenceDirectory).Path 'excel-spot-check.pdf'))
    $report = [ordered]@{
        passed = $true
        application = 'Microsoft Excel'
        version = $excel.Version
        build = $excel.Build
        workbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
        names = @($workbook.Worksheets | ForEach-Object { $_.Name })
        literal = $sheet.Cells.Item(4, 1).Value2
        unicode = $sheet.Cells.Item(2, 1).Value2
        dateFormat = $sheet.Cells.Item(2, 3).NumberFormat
        width = $originalWidth
        frozenHeader = $true
        dateText = $dateText
        pdfPreviewAutoFit = $true
        autoFilter = $sheet.AutoFilter.Range.Address()
        headerFill = $sheet.Cells.Item(1, 1).Interior.Color
    }
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'excel-spot-check.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 5
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
}
