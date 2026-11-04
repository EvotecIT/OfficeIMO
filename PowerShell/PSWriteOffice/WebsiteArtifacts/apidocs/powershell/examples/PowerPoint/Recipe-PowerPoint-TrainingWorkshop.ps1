param(
    [string] $OutputDirectory = (Join-Path $PWD 'PowerPoint-TrainingWorkshop')
)

Import-Module PSWriteOffice -ErrorAction Stop
$null = New-Item -Path $OutputDirectory -ItemType Directory -Force
$path = Join-Path $OutputDirectory 'Training-Workshop.pptx'
$agenda = @(
    [pscustomobject]@{ Module = 'Understand'; Duration = '15 min'; Outcome = 'Explain the operating model' }
    [pscustomobject]@{ Module = 'Practice'; Duration = '30 min'; Outcome = 'Complete the guided exercise' }
    [pscustomobject]@{ Module = 'Apply'; Duration = '20 min'; Outcome = 'Plan the first production use' }
)

New-OfficePowerPoint -Path $path -Theme Modern -Force -PassThru {
    PowerPointSlide -Title 'Automation Workshop' -Layout Title -Subtitle 'From repeatable data to reviewable documents' {
        PowerPointNotes 'Ask participants to name one document they rebuild manually every week.'
    }
    PowerPointSlide -Title 'Learning objectives' {
        PowerPointList -Items 'Choose the right document format', 'Compose content with a PowerShell DSL', 'Validate the generated artifact', 'Keep data and presentation concerns separate' `
            -X 2cm -Y 4cm -Width 28cm -Height 6cm -BulletChar '•'
        PowerPointNotes 'Connect every objective to the participant examples collected at the start.'
    }
    PowerPointSlide -Title 'Workshop agenda' {
        PowerPointTable -InputObject $agenda
        PowerPointNotes 'Take questions after each module rather than saving them for the end.'
    }
    PowerPointSlide -Title 'Your next step' {
        PowerPointCard 'First production report' -Text 'Pick one real report, replace the sample data, and review the generated file with its owner.'
        PowerPointNotes 'End with a concrete commitment from each participant.'
    }
}
