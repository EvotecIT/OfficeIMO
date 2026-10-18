Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot 'Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'ExamplePowerPoint6.pptx'
$presentation = New-OfficePowerPoint -Path $path -NoSave
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1
Add-OfficePowerPointSlide -Presentation $presentation -Layout 1

Remove-OfficePowerPointSlide -Presentation $presentation -Index 0
Save-OfficePowerPoint -Presentation $presentation

Write-Host "Presentation saved to $path"
