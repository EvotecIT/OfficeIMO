Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'PowerPoint-Charts.pptx'
$rows = @(
    [PSCustomObject]@{ Month = 'Jan'; MonthNumber = 1; Sales = 10; Profit = 4 }
    [PSCustomObject]@{ Month = 'Feb'; MonthNumber = 2; Sales = 14; Profit = 6 }
    [PSCustomObject]@{ Month = 'Mar'; MonthNumber = 3; Sales = 18; Profit = 8 }
)

$ppt = New-OfficePowerPoint -Path $path

$columnSlide = Add-OfficePowerPointSlide -Document $ppt -PassThru
Set-OfficePowerPointSlide -Slide $columnSlide -Title 'Column Chart'
Add-OfficePowerPointChart -Slide $columnSlide -Data $rows -Category Month -Value Sales, Profit -Title 'Sales vs Profit'

$pieSlide = Add-OfficePowerPointSlide -Document $ppt -PassThru
Set-OfficePowerPointSlide -Slide $pieSlide -Title 'Pie Chart'
Add-OfficePowerPointChart -Slide $pieSlide -Type Pie -Data $rows -Category Month -Value Sales -Title 'Sales Mix'

$scatterSlide = Add-OfficePowerPointSlide -Document $ppt -PassThru
Set-OfficePowerPointSlide -Slide $scatterSlide -Title 'Scatter Chart'
Add-OfficePowerPointChart -Slide $scatterSlide -Type Scatter -Data $rows -XProperty MonthNumber -YProperty Sales, Profit -Title 'Trend Scatter'

Save-OfficePowerPoint -Document $ppt
$ppt | Close-OfficePowerPoint

Write-Host "Presentation saved to $path"
