Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'ExamplePowerPoint1.pptx'
$presentation = New-OfficePowerPoint -Path $path -NoSave

$slide1 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide1 -Title 'Status Update'
Add-OfficePowerPointTextBox -Slide $slide1 -Text 'Generated with PSWriteOffice' -X 80 -Y 150 -Width 320 -Height 40
Add-OfficePowerPointShape -Slide $slide1 -ShapeType Rectangle -X 80 -Y 210 -Width 320 -Height 120 -FillColor '#DDEEFF' -OutlineColor '#4472C4' -OutlineWidth 1

$slide2 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide2 -Title 'Next Steps'
Add-OfficePowerPointTextBox -Slide $slide2 -Text '1. Review numbers  2. Plan Q1  3. Ship' -X 80 -Y 150 -Width 360 -Height 80

Save-OfficePowerPoint -Presentation $presentation
Write-Host "Presentation saved to $path"
