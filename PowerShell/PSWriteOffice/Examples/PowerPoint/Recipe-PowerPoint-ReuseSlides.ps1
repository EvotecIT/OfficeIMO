$sourcePath = '.\PowerPoint-Reusable-Slides.pptx'
$targetPath = '.\PowerPoint-Combined-Deck.pptx'

PptNew -Path $sourcePath {
    PptSlide {
        PptTitle -Title 'Reusable Architecture'
        PowerPointTextBox -Text 'Shared platform diagram' -X 80 -Y 150 -Width 500 -Height 80
    }
}

PptNew -Path $targetPath {
    PptSlide {
        PptTitle -Title 'Customer Briefing'
        PowerPointTextBox -Text 'Prepared for review' -X 80 -Y 150 -Width 500 -Height 80
    }
}

$target = Get-OfficePowerPoint -Path $targetPath
Import-OfficePowerPointSlide -Presentation $target -SourcePath $sourcePath -SourceIndex 0 -InsertAt 1
Copy-OfficePowerPointSlide -Presentation $target -Index 0 -InsertAt 2
Add-OfficePowerPointSection -Document $target -Name 'Shared material' -Slide 2
Close-OfficePowerPoint -Presentation $target -Save
