Import-Module PSParseHTML -ErrorAction Stop
Import-Module PSWriteOffice -ErrorAction Stop

$outputDirectory = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -ItemType Directory -Path $outputDirectory -Force
$htmlPath = Join-Path $outputDirectory 'HtmlTables.html'
$excelPath = Join-Path $outputDirectory 'HtmlTables.xlsx'

@'
<table id="results">
  <caption>Results</caption>
  <thead><tr><th>Name</th><th>Status</th></tr></thead>
  <tbody>
    <tr><td><a href="https://example.com/a">Alpha</a></td><td>Ready</td></tr>
    <tr><td><a href="https://example.com/b">Beta</a></td><td>Hold</td></tr>
  </tbody>
</table>
'@ | Set-Content -LiteralPath $htmlPath -Encoding UTF8

ConvertFrom-HtmlTable -Path $htmlPath -TableId 'results' -AsDataTable -IncludeLinkUrls |
    Export-OfficeExcel -Path $excelPath -WorksheetName 'Results' -TableName 'Results' -AutoFit -FreezeTopRow -BoldTopRow

$allTables = ConvertFrom-HtmlTable -Path $htmlPath -AsDataSet -IncludeLinkUrls
$allTables | Export-OfficeExcel -Path $excelPath -Append -AutoFit
