$path = '.\PowerPoint-Object-Composition.pptx'
$presentation = New-OfficePowerPoint -Path $path -NoSave

$titleSlide = Add-OfficePowerPointSlide -Presentation $presentation -LayoutType Title -PassThru
Set-OfficePowerPointSlideTitle -Slide $titleSlide -Title 'Customer onboarding review'
Add-OfficePowerPointTextBox -Slide $titleSlide -Text 'Decisions, owners, and the next milestone' -X 90 -Y 190 -Width 700 -Height 70
Set-OfficePowerPointNotes -Slide $titleSlide -Text 'Open with the customer outcome, then confirm the two decisions.'

$actionSlide = Add-OfficePowerPointSlide -Presentation $presentation -LayoutType Text -PassThru
Set-OfficePowerPointSlideTitle -Slide $actionSlide -Title 'Actions'
Add-OfficePowerPointTextBox -Slide $actionSlide -Run @{
    Text = 'Owner: ', 'Delivery', '    Due: ', 'Friday'
    Bold = $true, $false, $true, $true
} -X 90 -Y 170 -Width 700 -Height 60

$presentation | Save-OfficePowerPoint
$presentation | Close-OfficePowerPoint
