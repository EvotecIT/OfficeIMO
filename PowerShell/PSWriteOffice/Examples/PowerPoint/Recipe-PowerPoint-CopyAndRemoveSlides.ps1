$path = '.\PowerPoint-Copy-And-Remove.pptx'

$presentation = New-OfficePowerPoint -Path $path -NoSave
$overview = Add-OfficePowerPointSlide -Presentation $presentation -PassThru
Set-OfficePowerPointSlideTitle -Slide $overview -Title 'Reusable overview'
Add-OfficePowerPointTextBox -Slide $overview -Text 'Shared service story' -X 80 -Y 150 -Width 560 -Height 70

Copy-OfficePowerPointSlide -Presentation $presentation -Index 0 -InsertAt 1
$copied = Get-OfficePowerPointSlide -Presentation $presentation -Index 1
Set-OfficePowerPointSlideTitle -Slide $copied -Title 'Customer-specific overview'

$draft = Add-OfficePowerPointSlide -Presentation $presentation -PassThru
Set-OfficePowerPointSlideTitle -Slide $draft -Title 'Draft slide to remove'
Remove-OfficePowerPointSlide -Presentation $presentation -Index 2

$presentation | Save-OfficePowerPoint
$presentation | Close-OfficePowerPoint
