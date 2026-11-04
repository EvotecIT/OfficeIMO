Import-Module PSWriteOffice -ErrorAction Stop

$outputDirectory = Join-Path $PSScriptRoot '..\Documents'
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    $null = New-Item -Path $outputDirectory -ItemType Directory -Force
}

$path = Join-Path $outputDirectory 'Excel-ImportExport.xlsx'

$rows = @(
    [PSCustomObject]@{ Region = 'NA'; Revenue = 120000; Owner = 'Avery' }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 185000; Owner = 'Morgan' }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 142000; Owner = 'Jordan' }
)

$rows |
    Export-OfficeExcel -Path $path -Sheet 'Revenue' -TableName 'Revenue' -Title 'Regional Revenue'    -PassThru |
    Format-List FullName

Import-OfficeExcel -Path $path -Sheet 'Revenue' -Range 'A2:C5' |
    Format-Table -AutoSize
