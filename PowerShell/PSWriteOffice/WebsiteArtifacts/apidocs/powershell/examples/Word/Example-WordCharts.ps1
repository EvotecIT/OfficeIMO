Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$rows = @(
    [PSCustomObject]@{ Region = 'North America'; Revenue = 125000 }
    [PSCustomObject]@{ Region = 'EMEA'; Revenue = 98000 }
    [PSCustomObject]@{ Region = 'APAC'; Revenue = 143000 }
)

$trend = @(
    [PSCustomObject]@{ Month = 'Jan'; Sales = 10; Profit = 4 }
    [PSCustomObject]@{ Month = 'Feb'; Sales = 12; Profit = 5 }
    [PSCustomObject]@{ Month = 'Mar'; Sales = 15; Profit = 7 }
)

$tableRows = @(
    [PSCustomObject]@{ Topic = 'Regional revenue'; Details = 'See chart' }
)

$docPath = Join-Path $documents 'Word-Charts.docx'

New-OfficeWord -Path $docPath {
    Add-OfficeWordParagraph -Text 'Pie chart from PowerShell objects'
    Add-OfficeWordChart -Type Pie -Data $rows -CategoryProperty Region -SeriesProperty Revenue -Title 'Regional Revenue Mix' -FitToPageWidth -WidthFraction 0.70

    Add-OfficeWordParagraph -Text 'Line chart with two series'
    Add-OfficeWordChart -Type Line -Data $trend -CategoryProperty Month -SeriesProperty Sales, Profit -Legend -XAxisTitle 'Month' -YAxisTitle 'Value' -SeriesColor '#1f77b4', '#ff7f0e'

    Add-OfficeWordParagraph -Text 'Pie chart anchored inside a table cell'
    Add-OfficeWordTable -InputObject $tableRows -Style 'GridTable1LightAccent1' {
        WordTableCell -Row 1 -Column 1 {
            WordParagraph {
                WordChart -Type Pie -Data $rows -CategoryProperty Region -SeriesProperty Revenue -Title 'Cell Revenue Mix' -WidthPixels 420 -HeightPixels 280
            }
        }
    }
}

Write-Host "Document saved to $docPath"
