$path = '.\PowerPoint-Sections-And-Notes.pptx'

$presentation = New-OfficePowerPoint -Path $path -NoSave
$cover = Add-OfficePowerPointSlide -Presentation $presentation -PassThru
Set-OfficePowerPointSlideTitle -Slide $cover -Title 'Service review'
Set-OfficePowerPointNotes -Slide $cover -Text 'Introduce the reporting period and desired decision.'

$evidence = Add-OfficePowerPointSlide -Presentation $presentation -PassThru
Set-OfficePowerPointSlideTitle -Slide $evidence -Title 'Evidence'
Add-OfficePowerPointTextBox -Slide $evidence -Text 'Availability remained above 99.9%.' -X 80 -Y 150 -Width 650 -Height 70
Set-OfficePowerPointNotes -Slide $evidence -Text 'Pause for questions before moving to actions.'

Add-OfficePowerPointSection -Presentation $presentation -Name 'Briefing' -StartSlideIndex 0
Add-OfficePowerPointSection -Presentation $presentation -Name 'Evidence' -StartSlideIndex 1
$presentation | Save-OfficePowerPoint
$presentation | Close-OfficePowerPoint
