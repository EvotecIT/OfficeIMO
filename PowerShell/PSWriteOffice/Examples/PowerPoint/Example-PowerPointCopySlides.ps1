Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'PowerPoint-CopySlides.pptx'
$presentation = New-OfficePowerPoint -Path $path -NoSave

$intro = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $intro -Title 'Executive Summary'
Add-OfficePowerPointTextBox -Slide $intro -Text 'Quarterly revenue and margin summary' -X 80 -Y 150 -Width 360 -Height 60
Set-OfficePowerPointNotes -Slide $intro -Text 'Use this for board prep.'

$closing = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $closing -Title 'Appendix'

Copy-OfficePowerPointSlide -Presentation $presentation -Index 0 -InsertAt 1

Save-OfficePowerPoint -Presentation $presentation

Write-Host "Presentation saved to $path"
