param(
    [string] $OutputDirectory = (Join-Path (Get-Location) 'PowerPoint-TransitionsAndSizing')
)

$path = Join-Path $OutputDirectory 'PowerPoint-TransitionsAndSizing.pptx'
$document = New-OfficePowerPoint
try {
    Set-OfficePowerPointSlideSize -Document $document -Preset Screen16x9

    $intro = PowerPointSlide -Document $document -Title 'Executive summary' -PassThru {
        PowerPointParagraph 'Review the outcome, then agree on the next steps.'
    }
    Set-OfficePowerPointSlide -Slide $intro -Transition Fade

    $details = PowerPointSlide -Document $document -Title 'Details' -PassThru {
        PowerPointList 'Confirm the priorities', 'Assign the actions', 'Set the next review date'
    }
    Set-OfficePowerPointSlide -Slide $details -Transition Morph

    Save-OfficePowerPoint -Document $document -Path $path -PassThru
} finally {
    Close-OfficePowerPoint -Document $document
}
