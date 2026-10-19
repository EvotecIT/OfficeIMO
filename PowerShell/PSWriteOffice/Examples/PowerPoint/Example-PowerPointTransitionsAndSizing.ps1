Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'PowerPoint-TransitionsAndSizing.pptx'

$presentation = New-OfficePowerPoint -Path $path -NoSave
Set-OfficePowerPointSlideSize -Presentation $presentation -Preset Screen16x9

$intro = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $intro -Title 'Executive Summary'
Set-OfficePowerPointSlideTransition -Slide $intro -Transition Fade

$details = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $details -Title 'Details'
Set-OfficePowerPointSlideTransition -Slide $details -Transition Morph

Save-OfficePowerPoint -Presentation $presentation

Write-Host "Presentation saved to $path"
