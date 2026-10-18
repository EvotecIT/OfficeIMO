Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'PowerPoint-Charts.pptx'
$rows = @(
    [PSCustomObject]@{ Month = 'Jan'; MonthNumber = 1; Sales = 10; Profit = 4 }
    [PSCustomObject]@{ Month = 'Feb'; MonthNumber = 2; Sales = 14; Profit = 6 }
    [PSCustomObject]@{ Month = 'Mar'; MonthNumber = 3; Sales = 18; Profit = 8 }
)

$ppt = New-OfficePowerPoint -Path $path -NoSave

$columnSlide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $columnSlide -Title 'Column Chart'
Add-OfficePowerPointChart -Slide $columnSlide -Data $rows -CategoryProperty Month -SeriesProperty Sales, Profit -Title 'Sales vs Profit'

$pieSlide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $pieSlide -Title 'Pie Chart'
Add-OfficePowerPointChart -Slide $pieSlide -Type Pie -Data $rows -CategoryProperty Month -SeriesProperty Sales -Title 'Sales Mix'

$scatterSlide = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $scatterSlide -Title 'Scatter Chart'
Add-OfficePowerPointChart -Slide $scatterSlide -Type Scatter -Data $rows -XProperty MonthNumber -YProperty Sales, Profit -Title 'Trend Scatter'

Save-OfficePowerPoint -Presentation $ppt
$ppt | Close-OfficePowerPoint

Write-Host "Presentation saved to $path"
