$path = Join-Path $PWD 'PowerPoint-Inspection-Source.pptx'
New-OfficePowerPoint -Path $path {
    PowerPointSlide -Title 'Service Review' {
        PowerPointParagraph 'Executive summary'
    }
    PowerPointSlide -Title 'Next Steps' {
        PowerPointList 'Assign owner', 'Confirm date'
    }
}

$document = Open-OfficePowerPoint -Path $path
try {
    $slide = Get-OfficePowerPointSlide -Document $document -Slide 1
    Set-OfficePowerPointNotes -Slide $slide -Text 'Lead with the decision.'
    Save-OfficePowerPoint -Document $document

    Get-OfficePowerPointSlide -Document $document -Summary |
        Select-Object SlideIndex, Title, ShapeCount, HasNotes, NotesText
} finally {
    Close-OfficePowerPoint -Document $document
}

# File inspection returns detached summaries and closes the source after reading it.
Get-Item -LiteralPath $path | Get-OfficePowerPointSlide -Summary |
    Select-Object Path, SlideIndex, Title, ShapeCount, HasNotes
