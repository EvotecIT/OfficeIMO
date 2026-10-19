Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Csv-Advanced.csv'
$rows = @(
    [PSCustomObject]@{ Name = 'Alpha'; Score = 92; Active = $true }
    [PSCustomObject]@{ Name = 'Beta'; Score = 76; Active = $true }
    [PSCustomObject]@{ Name = 'Gamma'; Score = 64; Active = $false }
)

$rows | Export-OfficeCsv -Path $path -Delimiter ';'

Write-Host "CSV saved to $path"
Import-OfficeCsv -Path $path -Delimiter ';' -AsHashtable | Format-Table
