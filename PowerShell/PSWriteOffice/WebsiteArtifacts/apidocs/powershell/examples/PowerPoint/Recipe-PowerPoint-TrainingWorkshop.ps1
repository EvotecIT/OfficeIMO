$path = '.\Training-Workshop.pptx'
$agenda = @(
    [pscustomobject]@{ Module = 'Understand'; Duration = '15 min'; Outcome = 'Explain the operating model' }
    [pscustomobject]@{ Module = 'Practice'; Duration = '30 min'; Outcome = 'Complete the guided exercise' }
    [pscustomobject]@{ Module = 'Apply'; Duration = '20 min'; Outcome = 'Plan the first production use' }
)

PptNew -Path $path {
    PptSlideSize -Preset Screen16x9

    PptSlide {
        PptTitle -Title 'Automation Workshop'
        PptTextBox -Text 'From repeatable data to reviewable documents' -X 90 -Y 190 -Width 720 -Height 70
        PptNotes -Text 'Ask participants to name one document they rebuild manually every week.'
    }

    PptSlide {
        PptTitle -Title 'Learning objectives'
        PptBullets -Bullets 'Choose the right document format', 'Compose content with a PowerShell DSL', 'Validate the generated artifact', 'Keep data and presentation concerns separate' -X 90 -Y 135 -Width 700 -Height 260
        PptNotes -Text 'Connect every objective to the participant examples collected at the start.'
    }

    PptSlide {
        PptTitle -Title 'Workshop agenda'
        PptTable -Data $agenda -X 70 -Y 135 -Width 680 -Height 230
        PptNotes -Text 'Take questions after each module rather than saving them for the end.'
    }

    PptSlide {
        PptTitle -Title 'Your next step'
        PptShape -ShapeType RoundRectangle -X 95 -Y 160 -Width 650 -Height 180 -FillColor '#DBEAFE' -OutlineColor '#2563EB'
        PptTextBox -Text 'Pick one real report, replace the sample data, and review the generated file with its owner.' -X 135 -Y 215 -Width 570 -Height 90
        PptNotes -Text 'End with a concrete commitment from each participant.'
    }
}
