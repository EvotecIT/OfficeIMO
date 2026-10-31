BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $manifest -Global -ErrorAction Stop
}

Describe 'PowerPoint content workflows through the module' {
    It 'saves collected speaker notes separately from the visible slide content' {
        $path = Join-Path $TestDrive 'Speaker-notes.pptx'
        New-OfficePowerPoint -Path $path {
            PowerPointSlide -Title 'Service review' {
                PowerPointColumn {
                    PowerPointParagraph 'All regions healthy'
                    PowerPointNotes 'Pause for questions about capacity.'
                }
            }
        }
        $document = Open-OfficePowerPoint -LiteralPath $path
        try {
            $slide = Get-OfficePowerPointSlide -Document $document -Slide 1
            (Get-OfficePowerPointNotes -Slide $slide).Text | Should -Be 'Pause for questions about capacity.'
            (($slide.TextBoxes | ForEach-Object Text) -join "`n") | Should -Not -Match 'Pause for questions'
            $edited = $slide | Set-OfficePowerPointNotes -Text 'Notes edited after opening.' -PassThru
            [object]::ReferenceEquals($slide, $edited) | Should -BeTrue
            Save-OfficePowerPoint -Document $document
        } finally {
            Close-OfficePowerPoint -Document $document
        }
        (Get-OfficePowerPointSlide -LiteralPath $path -Summary).NotesText | Should -Be 'Notes edited after opening.'
    }
    It 'runs the object composition recipe and reopens its editable slides' {
        Push-Location $TestDrive
        try {
            & (Join-Path $PSScriptRoot '..\Examples\PowerPoint\Recipe-PowerPoint-ObjectComposition.ps1')
            $summary = @(Get-OfficePowerPointSlide -LiteralPath (Join-Path $TestDrive 'PowerPoint-Object-Composition.pptx') -Summary)
            $summary.Count | Should -Be 2
            $summary[0].Title | Should -Be 'Customer onboarding review'
            $summary[0].HasNotes | Should -BeTrue
            $summary[0].Slide | Should -BeNullOrEmpty
            $document = Open-OfficePowerPoint -LiteralPath (Join-Path $TestDrive 'PowerPoint-Object-Composition.pptx') -ErrorAction Stop
            try {
                $document.Slides.Count | Should -Be 2
                (Get-OfficePowerPointNotes -Document $document -Slide 1).Text | Should -Be 'Open with the customer outcome, then confirm the two decisions.'
                $text = ($document.Slides | ForEach-Object { $_.TextBoxes | ForEach-Object Text }) -join "`n"
                $text | Should -Match 'Customer onboarding review'
                $text | Should -Match 'Record the handover'
                @($document.ValidateDocument()).Count | Should -Be 0
            } finally {
                Close-OfficePowerPoint -Document $document
            }
        } finally {
            Pop-Location
        }
    }
    It 'runs the inspection recipe with one-based slides and detached file summaries' {
        Push-Location $TestDrive
        try {
            $result = @(& (Join-Path $PSScriptRoot '..\Examples\PowerPoint\Recipe-PowerPoint-InspectDeck.ps1'))
            $result.Count | Should -Be 4
            $result[0].SlideIndex | Should -Be 1
            $result[0].Title | Should -Be 'Service Review'
            $result[0].HasNotes | Should -BeTrue
            $result[0].NotesText | Should -Be 'Lead with the decision.'
            $result[1].SlideIndex | Should -Be 2
            $result[1].Title | Should -Be 'Next Steps'
            $result[2].Path | Should -Be (Join-Path $TestDrive 'PowerPoint-Inspection-Source.pptx')
            $result[2].HasNotes | Should -BeTrue
            $result[3].SlideIndex | Should -Be 2
            $file = [System.IO.File]::Open((Join-Path $TestDrive 'PowerPoint-Inspection-Source.pptx'), 'Open', 'ReadWrite', 'None')
            $file.Dispose()
        } finally {
            Pop-Location
        }
    }
    It 'edits selected slides through the consolidated native command and saves the result' {
        $path = Join-Path $TestDrive 'Edited-Slides.pptx'
        $document = New-OfficePowerPoint
        try {
            $first = Add-OfficePowerPointSlide -Document $document -Title 'First' -PassThru
            Add-OfficePowerPointSlide -Document $document -Title 'Second'
            $edited = Set-OfficePowerPointSlide 1 -Document $document -Title 'Updated' -Background Accent2 -Transition Fade -PassThru
            [object]::ReferenceEquals($first, $edited) | Should -BeTrue
            (Get-OfficePowerPointSlide -Document $document -Slide 1 -Summary).Title | Should -Be 'Updated'
            (Get-OfficePowerPointSlide -Document $document -Slide 2 -Summary).Title | Should -Be 'Second'
            $edited.Transition | Should -Be 'Fade'
            Set-OfficePowerPointSlide -Slide $edited -ClearBackground
            Save-OfficePowerPoint -Document $document -Path $path
        } finally {
            Close-OfficePowerPoint -Document $document
        }
        $saved = @(Get-OfficePowerPointSlide -LiteralPath $path -Summary)
        $saved.Count | Should -Be 2
        $saved[0].Title | Should -Be 'Updated'
        $saved[1].Title | Should -Be 'Second'
    }
    It 'runs the themed background and column example with the current commands' {
        $directory = Join-Path $TestDrive 'Backgrounds'
        $file = & (Join-Path $PSScriptRoot '..\Examples\PowerPoint\Example-PowerPointBackgroundsAndLayout.ps1') -OutputDirectory $directory
        $file | Should -BeOfType ([System.IO.FileInfo])
        $document = Open-OfficePowerPoint -LiteralPath $file.FullName
        try {
            $document.Slides.Count | Should -Be 2
            $document.Slides[0].GetBackground().Color | Should -Be 'F4F7FB'
            $document.Slides[1].GetBackground().Kind | Should -Be 'Image'
            $document.Slides[1].Transition | Should -Be 'Fade'
            $text = ($document.Slides[0].TextBoxes | ForEach-Object Text) -join "`n"
            $text | Should -Match 'Column 1'
            $text | Should -Match 'Column 2'
            @($document.ValidateDocument()).Count | Should -Be 0
        } finally {
            Close-OfficePowerPoint -Document $document
        }
    }
}
