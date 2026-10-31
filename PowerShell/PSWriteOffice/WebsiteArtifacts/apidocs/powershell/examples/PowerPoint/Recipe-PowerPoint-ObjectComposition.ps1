$path = '.\PowerPoint-Object-Composition.pptx'
$document = New-OfficePowerPoint -Theme Modern
try {
    $cover = Add-OfficePowerPointSlide -Document $document -Title 'Customer onboarding review' `
        -Subtitle 'Decisions, owners, and the next milestone' -PassThru
    Set-OfficePowerPointNotes -Slide $cover -Text 'Open with the customer outcome, then confirm the two decisions.'

    Add-OfficePowerPointSlide -Document $document -Title 'Actions' {
        PowerPointParagraph 'Delivery owns the next milestone, due Friday.'
        PowerPointList 'Confirm the acceptance criteria', 'Review progress with the customer', 'Record the handover'
    }

    Save-OfficePowerPoint -Document $document -Path $path
} finally {
    Close-OfficePowerPoint -Document $document
}
