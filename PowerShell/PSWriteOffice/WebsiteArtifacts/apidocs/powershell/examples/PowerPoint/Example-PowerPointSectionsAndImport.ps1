Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$sourcePath = Join-Path $documents 'PowerPoint-Source.pptx'
$targetPath = Join-Path $documents 'PowerPoint-SectionsAndImport.pptx'

$source = New-OfficePowerPoint -Path $sourcePath -NoSave
$sourceSlide = Add-OfficePowerPointSlide -Presentation $source -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $sourceSlide -Title 'FY24 Imported'
Add-OfficePowerPointTextBox -Slide $sourceSlide -Text 'FY24 details from source deck' -X 80 -Y 150 -Width 320 -Height 60
Set-OfficePowerPointNotes -Slide $sourceSlide -Text 'FY24 source notes'
Save-OfficePowerPoint -Presentation $source

$target = New-OfficePowerPoint -Path $targetPath -NoSave
$slide1 = Add-OfficePowerPointSlide -Presentation $target -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide1 -Title 'FY24 Overview'
Add-OfficePowerPointTextBox -Slide $slide1 -Text 'FY24 summary for leadership' -X 80 -Y 150 -Width 320 -Height 60

$slide2 = Add-OfficePowerPointSlide -Presentation $target -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide2 -Title 'FY24 Results'

Add-OfficePowerPointSection -Presentation $target -Name 'Intro' -StartSlideIndex 0
Add-OfficePowerPointSection -Presentation $target -Name 'Results' -StartSlideIndex 1
Rename-OfficePowerPointSection -Presentation $target -Name 'Results' -NewName 'Deep Dive'

Update-OfficePowerPointText -Presentation $target -OldValue 'FY24' -NewValue 'FY25'
Copy-OfficePowerPointSlide -Presentation $target -Index 0 -InsertAt 1
Import-OfficePowerPointSlide -Presentation $target -SourcePath $sourcePath -SourceIndex 0 -InsertAt 1

Save-OfficePowerPoint -Presentation $target

Write-Host "Target deck saved to $targetPath"
Write-Host ''
Write-Host 'Sections:'
$reloaded = Get-OfficePowerPoint -Path $targetPath
try {
    Get-OfficePowerPointSection -Presentation $reloaded | Format-Table
} finally {
    $reloaded | Close-OfficePowerPoint
}
