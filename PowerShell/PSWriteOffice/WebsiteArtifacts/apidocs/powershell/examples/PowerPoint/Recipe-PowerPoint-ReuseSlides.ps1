$sourcePath = '.\PowerPoint-Reusable-Slides.pptx'
$targetPath = '.\PowerPoint-Combined-Deck.pptx'

PptNew -Path $sourcePath {
    PptSlide {
        PptTitle -Title 'Reusable Architecture'
        PptTextBox -Text 'Shared platform diagram' -X 80 -Y 150 -Width 500 -Height 80
    }
}

PptNew -Path $targetPath {
    PptSlide {
        PptTitle -Title 'Customer Briefing'
        PptTextBox -Text 'Prepared for review' -X 80 -Y 150 -Width 500 -Height 80
    }
}

$target = Get-OfficePowerPoint -Path $targetPath
Import-OfficePowerPointSlide -Presentation $target -SourcePath $sourcePath -SourceIndex 0 -InsertAt 1
Copy-OfficePowerPointSlide -Presentation $target -Index 0 -InsertAt 2
Add-OfficePowerPointSection -Presentation $target -Name 'Shared material' -StartSlideIndex 1
Close-OfficePowerPoint -Presentation $target -Save
