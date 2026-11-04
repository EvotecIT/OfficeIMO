Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'LoadExample.pptx'
$presentation = New-OfficePowerPoint -Path $path
Add-OfficePowerPointSlide -Document $presentation -TemplateLayout 1
Save-OfficePowerPoint -Document $presentation

$loaded = Open-OfficePowerPoint -Path $path
Write-Host "Loaded presentation with $($loaded.Slides.Count) slide(s)."
