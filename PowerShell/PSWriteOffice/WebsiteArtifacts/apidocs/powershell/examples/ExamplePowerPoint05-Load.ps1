Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'LoadExample.pptx'
$presentation = New-OfficePowerPoint -Path $path -NoSave
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Save-OfficePowerPoint -Presentation $presentation

$loaded = Get-OfficePowerPoint -Path $path
Write-Host "Loaded presentation with $($loaded.Slides.Count) slide(s)."
