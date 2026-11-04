$path = '.\PowerPoint-Copy-And-Remove.pptx'

$presentation = New-OfficePowerPoint -Path $path
$overview = Add-OfficePowerPointSlide -Document $presentation -PassThru
Set-OfficePowerPointSlide -Slide $overview -Title 'Reusable overview'
Add-OfficePowerPointTextBox -Slide $overview -Text 'Shared service story' -X 80 -Y 150 -Width 560 -Height 70

Copy-OfficePowerPointSlide -Document $presentation -Slide 1 -Position 2
$copied = Get-OfficePowerPointSlide -Document $presentation -Slide 2
Set-OfficePowerPointSlide -Slide $copied -Title 'Customer-specific overview'

$draft = Add-OfficePowerPointSlide -Document $presentation -PassThru
Set-OfficePowerPointSlide -Slide $draft -Title 'Draft slide to remove'
Remove-OfficePowerPointSlide -Document $presentation -Slide 3

$presentation | Save-OfficePowerPoint
$presentation | Close-OfficePowerPoint
