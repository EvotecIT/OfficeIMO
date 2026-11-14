Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Data.csv'
$rows = @(
    [PSCustomObject]@{ Name = 'Alpha'; Value = 1 }
    [PSCustomObject]@{ Name = 'Beta'; Value = 2 }
)

$rows | Export-OfficeCsv -Path $path

$data = Import-OfficeCsv -Path $path
$data | Format-Table

# CSV text, such as ConvertTo-OfficeCsv output or lines from another tool, reads back the same way.
$text = $rows | ConvertTo-OfficeCsv
$text | ConvertFrom-OfficeCsv | Format-Table
"Name;Value`nGamma;3" | ConvertFrom-OfficeCsv -Delimiter ';' -As Hashtable
