Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'ExamplePowerPoint4.pptx'
$presentation = New-OfficePowerPoint -Path $path

$slide = Add-OfficePowerPointSlide -Document $presentation -TemplateLayout 1 -PassThru
Set-OfficePowerPointSlide -Slide $slide -Title 'Quarterly Report'
Add-OfficePowerPointTextBox -Slide $slide -Text 'Generated with PSWriteOffice' -X 90 -Y 160 -Width 320 -Height 50

Save-OfficePowerPoint -Document $presentation
Write-Host "Presentation saved to $path"
