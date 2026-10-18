Import-Module PSWriteOffice -ErrorAction Stop

$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$presentationPath = Join-Path $documents 'PowerPoint-HtmlReview.pptx'
$semanticHtmlPath = Join-Path $documents 'PowerPoint-HtmlReview.semantic.html'
$visualHtmlPath = Join-Path $documents 'PowerPoint-HtmlReview.visual.html'

$presentation = New-OfficePowerPoint -Path $presentationPath -NoSave

$statusSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $statusSlide -Title 'Monthly Service Review'
Add-OfficePowerPointTextBox -Slide $statusSlide -Text 'Identity, Messaging, and Reporting are ready for leadership review.' -X 80 -Y 140 -Width 560 -Height 80
Set-OfficePowerPointNotes -Slide $statusSlide -Text 'Use this slide to introduce the operational status summary.'

$tableSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $tableSlide -Title 'Open Items'
Add-OfficePowerPointTable -Slide $tableSlide -Headers 'Area', 'Owner', 'Next Step' -Rows @(
    @('Messaging', 'Collaboration', 'Review retry spikes')
    @('Reporting', 'Analytics', 'Publish refreshed dashboard')
) -X 70 -Y 130 -Width 600 -Height 160

Save-OfficePowerPoint -Presentation $presentation
$presentation | Close-OfficePowerPoint

ConvertTo-OfficePowerPointHtml -Path $presentationPath -OutputPath $semanticHtmlPath -Title 'Deck Review'
ConvertTo-OfficePowerPointHtml -Path $presentationPath -Profile VisualReview -OutputPath $visualHtmlPath -Title 'Deck Visual Review'

Write-Host "Presentation saved to $presentationPath"
Write-Host "Semantic HTML saved to $semanticHtmlPath"
Write-Host "Visual HTML saved to $visualHtmlPath"
