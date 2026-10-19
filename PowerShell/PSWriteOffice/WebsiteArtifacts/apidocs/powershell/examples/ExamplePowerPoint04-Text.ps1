Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'ExamplePowerPoint4.pptx'
$presentation = New-OfficePowerPoint -Path $path -NoSave

$slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Quarterly Report'
Add-OfficePowerPointTextBox -Slide $slide -Text 'Generated with PSWriteOffice' -X 90 -Y 160 -Width 320 -Height 50

Save-OfficePowerPoint -Presentation $presentation
Write-Host "Presentation saved to $path"
