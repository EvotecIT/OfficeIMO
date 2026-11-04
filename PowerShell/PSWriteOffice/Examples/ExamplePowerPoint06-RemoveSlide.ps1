Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'ExamplePowerPoint6.pptx'
$presentation = New-OfficePowerPoint -Path $path
Add-OfficePowerPointSlide -Document $presentation -TemplateLayout 1
Add-OfficePowerPointSlide -Document $presentation -TemplateLayout 1

Remove-OfficePowerPointSlide -Document $presentation -Slide 1
Save-OfficePowerPoint -Document $presentation

Write-Host "Presentation saved to $path"
