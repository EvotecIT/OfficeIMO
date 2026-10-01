param([string] $OutputDirectory = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}
if (-not (Test-Path -LiteralPath $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory | Out-Null }
$app = $null; $books = $null; $book = $null; $owned = $false
$results = @()
$formats = @('yyyy-mm-dd hh:mm:ss', '[h]:mm:ss', 'hh:mm:ss', 'hhmmss', 'mm:ss', 'mmhh', 'General')
$serials = @(-0.5, 0, 1, 1.5, 59, 59.5, 60, 60.5, 61, 61.5, 1462, 45000)
try {
    $app = New-Object -ComObject Excel.Application
    $books = $app.Workbooks
    if ($books.Count -ne 0) { throw 'The Excel instance contains existing workbooks.' }
    $owned = $true
    $app.Visible = $false; $app.DisplayAlerts = $false; $app.AutomationSecurity = 3
    # Use Excel's own localized format symbols; COM dispatch can select a localized LCID.
    $yearCode = [string]$app.International(19); $dayCode = [string]$app.International(21)
    $hourCode = [string]$app.International(22); $minuteCode = [string]$app.International(23)
    $secondCode = [string]$app.International(24)
    foreach ($date1904 in @($false, $true)) {
        $book = $books.Add()
        $book.Date1904 = $date1904
        $sheets = $null; $sheet = $null; $cell = $null; $column = $null
        try {
            $sheets = $book.Worksheets; $sheet = $sheets.Item(1); $sheet.Name = 'Data'
            for ($c = 1; $c -le $formats.Count; $c++) {
                $cell = $sheet.Cells.Item(1, $c); $cell.Value2 = "Field$c"; Release-Com $cell; $cell = $null
                $column = $sheet.Columns.Item($c)
                $localFormat = if ($formats[$c - 1] -eq 'General') { [string]$app.International(26) }
                    else { $formats[$c - 1].Replace('y', $yearCode).Replace('d', $dayCode).Replace('h', $hourCode).Replace('m', $minuteCode).Replace('s', $secondCode) }
                try { $column.NumberFormatLocal = $localFormat } catch { throw "Excel rejected column $c format '$localFormat': $($_.Exception.Message)" }
                $column.ColumnWidth = 35
                Release-Com $column; $column = $null
                for ($r = 0; $r -lt $serials.Count; $r++) {
                    $cell = $sheet.Cells.Item($r + 2, $c); $cell.Value2 = [double]$serials[$r]
                    Release-Com $cell; $cell = $null
                }
                $cell = $sheet.Cells.Item($serials.Count + 2, $c); $cell.Formula = '=1+0.5'
                Release-Com $cell; $cell = $null
            }
            $app.CalculateFull()
            foreach ($format in @(@{ Extension = 'xlsx'; Id = 51 }, @{ Extension = 'xls'; Id = 56 }, @{ Extension = 'xlsb'; Id = 50 })) {
                $name = 'serials-' + $(if ($date1904) { '1904' } else { '1900' }) + '.' + $format.Extension
                $path = Join-Path $OutputDirectory $name
                $book.SaveAs($path, $format.Id)
                $records = @()
                for ($r = 2; $r -le $serials.Count + 2; $r++) {
                    for ($c = 1; $c -le $formats.Count; $c++) {
                        $cell = $sheet.Cells.Item($r, $c)
                        $expected = if ($r -eq $serials.Count + 2) { 1.5 } else { $serials[$r - 2] }
                        if ($cell.Value2 -ne $expected) { throw "$name changed serial at $r/$c." }
                        $records += @{ Row = $r; Column = $c; Serial = $cell.Value2; RequestedFormat = $formats[$c - 1]; Format = $cell.NumberFormat; NativeText = $cell.Text; Formula = $cell.Formula }
                        Release-Com $cell; $cell = $null
                    }
                }
                $results += @{ File = $name; Date1904 = $book.Date1904; Cells = $records }
            }
        } finally {
            Release-Com $column; Release-Com $cell; Release-Com $sheet; Release-Com $sheets
            $book.Close($false); Release-Com $book; $book = $null
        }
    }
    foreach ($result in $results) { $result.Sha256 = (Get-FileHash -LiteralPath (Join-Path $OutputDirectory $result.File) -Algorithm SHA256).Hash }
    [ordered]@{ Producer = 'Microsoft Excel'; Version = $app.Version; Build = $app.Build; GeneratedUtc = [DateTime]::UtcNow.ToString('o'); Results = $results } |
        ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'provenance.json') -Encoding UTF8
} finally {
    if ($null -ne $book) { $book.Close($false); Release-Com $book }
    Release-Com $books
    if ($null -ne $app -and $owned) { $app.Quit() }
    Release-Com $app
}
