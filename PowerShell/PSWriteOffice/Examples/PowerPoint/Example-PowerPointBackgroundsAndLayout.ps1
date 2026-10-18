Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'PowerPoint-BackgroundsAndLayout.pptx'
$imagePath = Join-Path $documents 'PowerPoint-Background.bmp'

[byte[]] $bytes = 0x42, 0x4D, 0x3A, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x36, 0x00, 0x00, 0x00, 0x28, 0x00,
    0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00,
    0x00, 0x00, 0x01, 0x00, 0x18, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x13, 0x0B,
    0x00, 0x00, 0x13, 0x0B, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF,
    0xFF, 0x00
[System.IO.File]::WriteAllBytes($imagePath, $bytes)

$ppt = New-OfficePowerPoint -Path $path -NoSave
Set-OfficePowerPointSlideSize -Presentation $ppt -WidthCm 30 -HeightCm 20

$slide1 = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide1 -Title 'Content Grid'
Set-OfficePowerPointBackground -Slide $slide1 -Color '#F4F7FB'

$columns = @(Get-OfficePowerPointLayoutBox -Presentation $ppt -ColumnCount 2 -MarginCm 1.5 -GutterCm 1.0)
foreach ($index in 0..($columns.Count - 1)) {
    $box = $columns[$index]
    Add-OfficePowerPointTextBox -Slide $slide1 -Text "Column $($index + 1)" -X $box.LeftPoints -Y $box.TopPoints -Width $box.WidthPoints -Height 48
}

$slide2 = Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru
Set-OfficePowerPointSlideTitle -Slide $slide2 -Title 'Image Background'
Set-OfficePowerPointBackground -Slide $slide2 -ImagePath $imagePath

Save-OfficePowerPoint -Presentation $ppt
$ppt | Close-OfficePowerPoint

Write-Host "Presentation saved to $path"
