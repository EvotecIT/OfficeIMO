$path = '.\Training-Workshop.pptx'
$agenda = @(
    [pscustomobject]@{ Module = 'Understand'; Duration = '15 min'; Outcome = 'Explain the operating model' }
    [pscustomobject]@{ Module = 'Practice'; Duration = '30 min'; Outcome = 'Complete the guided exercise' }
    [pscustomobject]@{ Module = 'Apply'; Duration = '20 min'; Outcome = 'Plan the first production use' }
)

New-OfficePowerPoint -Path $path -Force {
    PowerPointSlide -Title 'Automation workshop' -Subtitle 'From repeatable data to reviewable documents' {
        PowerPointNotes 'Ask participants to name one document they rebuild manually every week.'
    }

    PowerPointSlide -Title 'Learning objectives' {
        PowerPointList 'Choose the right document format', 'Compose content with a PowerShell DSL', 'Validate the generated artifact', 'Keep data and presentation concerns separate'
        PowerPointNotes 'Connect every objective to the participant examples collected at the start.'
    }

    # PowerPointStep items give the Process layout: one step per module.
    PowerPointSlide -Title 'Workshop agenda' {
        foreach ($module in $agenda) { PowerPointStep "$($module.Module), $($module.Duration)" $module.Outcome }
        PowerPointNotes 'Take questions after each module rather than saving them for the end.'
    }

    PowerPointSlide -Layout Closing -Title 'Your next step' -Subtitle 'Pick one real report, replace the sample data, and review the generated file with its owner.' {
        PowerPointNotes 'End with a concrete commitment from each participant.'
    }
}
