param(
    [string] $OutputDirectory = (Join-Path $PWD 'PowerPoint-BackgroundsAndLayout'),
    [string] $BackgroundImage
)

if (-not (Get-Module PSWriteOffice)) { Import-Module PSWriteOffice -ErrorAction Stop }
$directory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$null = [System.IO.Directory]::CreateDirectory($directory)
$path = Join-Path $directory 'PowerPoint-BackgroundsAndLayout.pptx'

if (-not $BackgroundImage) {
    # A small embedded image keeps the example runnable; supply a photo with -BackgroundImage.
    $BackgroundImage = Join-Path $directory 'PowerPoint-Background.bmp'
    [byte[]] $bytes = 0x42, 0x4D, 0x3A, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x36, 0x00, 0x00, 0x00, 0x28, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x18, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x13, 0x0B,
        0x00, 0x00, 0x13, 0x0B, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF,
        0xFF, 0x00
    [System.IO.File]::WriteAllBytes($BackgroundImage, $bytes)
}

$document = New-OfficePowerPoint -Theme Modern
try {
    $grid = Add-OfficePowerPointSlide -Document $document -Title 'Content grid' -PassThru {
        PowerPointColumn -Title 'Column 1' { PowerPointParagraph 'The layout engine places this content.' }
        PowerPointColumn -Title 'Column 2' { PowerPointParagraph 'Both columns share the document theme.' }
    }
    Set-OfficePowerPointSlide -Slide $grid -Background '#F4F7FB'

    $imageSlide = Add-OfficePowerPointSlide -Document $document -Title 'Image background' -PassThru
    Set-OfficePowerPointSlide -Slide $imageSlide -BackgroundImage $BackgroundImage -Transition Fade
    Save-OfficePowerPoint -Document $document -Path $path
} finally {
    Close-OfficePowerPoint -Document $document
}
Get-Item -LiteralPath $path
