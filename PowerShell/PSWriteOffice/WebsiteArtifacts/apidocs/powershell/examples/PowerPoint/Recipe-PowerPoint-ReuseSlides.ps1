param(
    [string] $OutputDirectory = (Join-Path $PWD 'PowerPoint-ReusedSlides')
)

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$sourcePath = Join-Path $OutputDirectory 'PowerPoint-Reusable-Slides.pptx'
$targetPath = Join-Path $OutputDirectory 'PowerPoint-Combined-Deck.pptx'

New-OfficePowerPoint -Path $sourcePath -Force {
    PowerPointSlide -Title 'Reusable Architecture' {
        PowerPointParagraph 'Shared platform diagram'
    }
}

New-OfficePowerPoint -Path $targetPath -Force -PassThru {
    PowerPointSlide -Title 'Customer Briefing' {
        PowerPointParagraph 'Prepared for review'
    }
    Import-OfficePowerPointSlide -LiteralPath $sourcePath -SourceSlide 1 -Position 2
    Copy-OfficePowerPointSlide -Slide 1 -Position 3
    Add-OfficePowerPointSection -Name 'Shared material' -Slide 2
}
