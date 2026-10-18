$path = '.\PowerPoint-Inspection-Source.pptx'
PptNew -Path $path {
    PptSlide {
        PptTitle -Title 'Service Review'
        PptTextBox -Text 'Executive summary' -X 80 -Y 150 -Width 420 -Height 60
        PptNotes -Text 'Lead with the decision.'
    }
    PptSlide {
        PptTitle -Title 'Next Steps'
        PptBullets -Bullets 'Assign owner', 'Confirm date' -X 80 -Y 150 -Width 420 -Height 160
    }
}

$presentation = Get-OfficePowerPoint -Path $path
foreach ($index in 0..($presentation.Slides.Count - 1)) {
    $slide = Get-OfficePowerPointSlide -Presentation $presentation -Index $index
    $summary = Get-OfficePowerPointSlideSummary -Slide $slide

    [pscustomobject]@{
        Slide  = $index + 1
        Title  = $summary.Title
        Shapes = @(Get-OfficePowerPointShape -Slide $slide).Count
        Notes  = @(Get-OfficePowerPointNotes -Slide $slide).Count
    }
}
Close-OfficePowerPoint -Presentation $presentation
