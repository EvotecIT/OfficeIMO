BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'PowerPointShapeEditing workflows' {
    BeforeEach {
        $presentation = $null
        $reloaded = $null
        $afterRemoval = $null
        $source = $null
        $target = $null
    }
    AfterEach {
        foreach ($document in @($presentation, $reloaded, $afterRemoval, $source, $target)) {
            if ($null -ne $document) {
                $document.Dispose()
            }
        }
    }

    It 'creates a presentation with shapes, tables, media, and notes' {
        $path = Join-Path $TestDrive 'PowerPointContent.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        $layouts = Get-OfficePowerPointLayout -Document $presentation
        $layouts.Count | Should -BeGreaterThan 0
        $selectedLayout = $layouts | Select-Object -First 1
        $layoutMaster = $selectedLayout.MasterIndex
        $layoutIndex = $selectedLayout.LayoutIndex
        $slide2 = $presentation.AddSlide($layoutMaster, $layoutIndex)

        $slide = $presentation.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $slide -Title 'Status Update'
        $shape = Add-OfficePowerPointShape -Slide $slide -ShapeType Rectangle -X 40 -Y 40 -Width 200 -Height 80 -FillColor '#DDEEFF' -OutlineColor '#1F4E79' -OutlineWidth 1 -PassThru

        $rows = @(
            [PSCustomObject]@{ Item = 'Alpha'; Qty = 10 }
            [PSCustomObject]@{ Item = 'Beta'; Qty = 20 }
        )
        $table = Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 40 -Y 140 -Width 360 -Height 200 -PassThru
        $image = Add-OfficePowerPointImage -Slide $slide -Path $imagePath -X 420 -Y 40 -Width 120 -Height 90 -PassThru
        $bullets = Add-OfficePowerPointBullets -Slide $slide -Bullets 'Wins','Risks','Next Steps' -X 420 -Y 150 -Width 250 -Height 200 -PassThru
        Set-OfficePowerPointNotes -Slide $slide -Text 'Keep this under five minutes.'

        $presentation.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointDocument'
        $slide.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointSlide'
        $shape.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointAutoShape'
        $table.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointTable'
        $image.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointPicture'
        $bullets.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointTextBox'
        $slide.Shapes.Count | Should -BeGreaterThan 0
        @($slide.Tables).Count | Should -BeGreaterThan 0
        $table.Rows | Should -BeGreaterThan 0
        $shape | Should -Not -BeNullOrEmpty
        $slide.Pictures.Count | Should -BeGreaterThan 0
        $image | Should -Not -BeNullOrEmpty
        $bullets.Paragraphs.Count | Should -Be 3
        $bullets.Paragraphs[0].Text | Should -Be 'Wins'
        $notes = Get-OfficePowerPointNotes -Slide $slide
        $notes.Text | Should -Be 'Keep this under five minutes.'
        $notes.HasNotes | Should -BeTrue
        $slideSummary = Get-OfficePowerPointSlide -Document $presentation -Slide ($slide.Index + 1) -Summary
        $slideSummary.Title | Should -Be 'Status Update'
        $slideSummary.HasNotes | Should -BeTrue
        $slideSummary.NotesText | Should -Be 'Keep this under five minutes.'
        $slideSummary.ShapeCount | Should -BeGreaterThan 0
        $slideSummary.PictureCount | Should -Be 1
        $slideSummary.TableCount | Should -Be 1
        $slideSummary.PlaceholderCount | Should -BeGreaterThan 0
        $slideSummary.LayoutPlaceholderCount | Should -BeGreaterThan 0
        $shapeInfo = Get-OfficePowerPointShape -Slide $slide
        $shapeInfo.Count | Should -BeGreaterThan 0
        $shapeInfo | Where-Object Kind -eq 'Picture' | Should -HaveCount 1
        $shapeInfo | Where-Object Kind -eq 'Table' | Should -HaveCount 1
        $shapeInfo | Where-Object Kind -eq 'AutoShape' | Should -HaveCount 1
        ($shapeInfo | Where-Object Kind -eq 'Picture' | Select-Object -First 1).MimeType | Should -Be 'image/bmp'
        ($shapeInfo | Where-Object Kind -eq 'Table' | Select-Object -First 1).RowCount | Should -Be 3
        ($shapeInfo | Where-Object Kind -eq 'Table' | Select-Object -First 1).ColumnCount | Should -Be 2
        $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
        $placeholder.Text | Should -Be 'Status Update'
        $placeholderUpdate = Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Title -Text 'Status Update v2' -PassThru
        $placeholderUpdate.Text | Should -Be 'Status Update v2'
        $layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
        $layoutPlaceholders.Count | Should -BeGreaterThan 0

        $slide2 | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Document $presentation

        Test-Path $path | Should -BeTrue
        $presentation.Dispose()
        $presentation = $null

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -Be 2

            $reloadedSlide = Get-OfficePowerPointSlide -Document $reloaded -Slide 2
            $reloadedPlaceholder = Get-OfficePowerPointPlaceholder -Slide $reloadedSlide -PlaceholderType Title
            if (-not $reloadedPlaceholder) {
                $reloadedPlaceholder = Get-OfficePowerPointPlaceholder -Slide $reloadedSlide -PlaceholderType CenteredTitle
            }
            $reloadedPlaceholder.Text | Should -Be 'Status Update v2'
            $reloadedNotes = Get-OfficePowerPointNotes -Slide $reloadedSlide
            $reloadedNotes.Text | Should -Be 'Keep this under five minutes.'
            $reloadedNotes.SlideIndex | Should -Be 2
            $reloadedSummary = Get-OfficePowerPointSlide -Document $reloaded -Slide 2 -Summary
            $reloadedSummary.Title | Should -Be 'Status Update v2'
            $reloadedSummary.SlideIndex | Should -Be 2
            $reloadedSummary.PictureCount | Should -Be 1
            $reloadedSummary.TableCount | Should -Be 1
            $reloadedShapeInfo = Get-OfficePowerPointShape -Presentation $reloaded -Index 1
            $reloadedShapeInfo.Count | Should -BeGreaterThan 0
            ($reloadedShapeInfo | Where-Object Kind -eq 'Picture').Count | Should -Be 1
            ($reloadedShapeInfo | Where-Object Kind -eq 'Table').Count | Should -Be 1
            ($reloadedShapeInfo | Where-Object Kind -eq 'Picture' | Select-Object -First 1).SlideIndex | Should -Be 1
            @($reloadedSlide.Tables).Count | Should -BeGreaterThan 0
            $reloadedSlide.Pictures.Count | Should -BeGreaterThan 0
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'finds existing PowerPoint shapes by metadata without a text term' {
        $path = Join-Path $TestDrive 'PowerPointMetadataShapeFind.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            $textBox = Add-OfficePowerPointTextBox -Slide $slide -Text 'Metadata status' -X 80 -Y 80 -Width 300 -Height 50 -PassThru
            $textBox.Name = 'Status.Primary'
            $rows = @(
                [PSCustomObject]@{ Metric = 'Risk'; State = 'Open' }
            )
            $table = Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 80 -Y 160 -Width 420 -Height 120 -PassThru
            $table.Name = 'Status.Table'

            $byName = Find-OfficePowerPointShape -Presentation $presentation -Name 'Status.*' -Kind TextBox
            $byName | Should -HaveCount 1
            $byName[0].Name | Should -Be 'Status.Primary'
            $byName[0].Kind | Should -Be 'TextBox'

            $byKind = Find-OfficePowerPointShape -Presentation $presentation -Kind Table
            $byKind | Should -HaveCount 1
            $byKind[0].Name | Should -Be 'Status.Table'
            $byKind[0].Kind | Should -Be 'Table'
        } finally {
            Close-OfficePowerPoint -Document $presentation
        }
    }

    It 'arranges PowerPoint shapes through OfficeIMO layout helpers' {
        $path = Join-Path $TestDrive 'PowerPointShapeLayout.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        try {
            $slide = $presentation.AddSlide(0, 1)
            $shape1 = Add-OfficePowerPointShape -Slide $slide -Name 'Kpi.One' -ShapeType Rectangle -X 40 -Y 80 -Width 80 -Height 40 -PassThru
            $shape2 = Add-OfficePowerPointShape -Slide $slide -Name 'Kpi.Two' -ShapeType Rectangle -X 180 -Y 140 -Width 80 -Height 40 -PassThru
            $shape3 = Add-OfficePowerPointShape -Slide $slide -Name 'Kpi.Three' -ShapeType Rectangle -X 320 -Y 200 -Width 80 -Height 40 -PassThru

            @($shape1, $shape2, $shape3) | Set-OfficePowerPointShapeLayout -Slide $slide -Align Top
            $shape1.Top.ToPoints() | Should -Be $shape2.Top.ToPoints()
            $shape1.Top.ToPoints() | Should -Be $shape3.Top.ToPoints()

            Find-OfficePowerPointShape -Slide $slide -Name 'Kpi.*' |
                Set-OfficePowerPointShapeLayout -Grid -Columns 3 -Rows 1 -ToSlide -GutterXPoints 12 -NoResize | Out-Null

            $arranged = @(Get-OfficePowerPointShape -Slide $slide -Kind AutoShape | Sort-Object LeftPoints)
            $arranged | Should -HaveCount 3
            $arranged[0].LeftPoints | Should -BeLessThan $arranged[1].LeftPoints
            $arranged[1].LeftPoints | Should -BeLessThan $arranged[2].LeftPoints
        } finally {
            Close-OfficePowerPoint -Document $presentation
        }
    }

    It 'reads notes without creating empty notes parts' {
        $path = Join-Path $TestDrive 'PowerPointNotesRead.pptx'
        $presentation = New-OfficePowerPoint -Path $path

        $slide = $presentation.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $slide -Title 'No notes yet'

        $notes = Get-OfficePowerPointNotes -Slide $slide -IncludeEmpty
        $notes.HasNotes | Should -BeFalse
        $notes.Text | Should -BeNullOrEmpty
        $notes.SlideIndex | Should -Be 1

        $presentationNotes = Get-OfficePowerPointNotes -Document $presentation -IncludeEmpty
        $presentationNotes.Count | Should -Be 1
        $presentationNotes[0].HasNotes | Should -BeFalse
        $slideSummary = Get-OfficePowerPointSlide -Document $presentation -Summary
        $slideSummary.Count | Should -Be 1
        $slideSummary[0].HasNotes | Should -BeFalse
        $slideSummary[0].Title | Should -Be 'No notes yet'

        $shapeInfo = Get-OfficePowerPointShape -Presentation $presentation -Index 0 -Kind TextBox
        $shapeInfo.Count | Should -BeGreaterThan 0
        $shapeInfo[0].SlideIndex | Should -Be 0
        $shapeByIndex = Get-OfficePowerPointShape -Slide $slide -ShapeIndex $shapeInfo[0].ShapeIndex
        $shapeByIndex.ShapeIndex | Should -Be $shapeInfo[0].ShapeIndex
        $shapeByIndex.Kind | Should -Be 'TextBox'

        Save-OfficePowerPoint -Document $presentation

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'ppt/notesSlides/*' }).Count | Should -Be 0
    }

    It 'persists layout placeholder edits across save and reopen' {
        $path = Join-Path $TestDrive 'PowerPointLayoutEdits.pptx'
        $presentation = New-OfficePowerPoint -Path $path
        $layoutPlaceholder = $null

        $layouts = Get-OfficePowerPointLayout -Document $presentation
        $layouts.Count | Should -BeGreaterThan 0

        $selectedLayout = $layouts | Select-Object -First 1
        $layoutMaster = $selectedLayout.MasterIndex
        $layoutIndex = $selectedLayout.LayoutIndex
        $slide = $presentation.AddSlide($layoutMaster, $layoutIndex)

        $layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
        $layoutPlaceholders.Count | Should -BeGreaterThan 0

        $layoutPlaceholder = $layoutPlaceholders | Where-Object { $_.PlaceholderType } | Select-Object -First 1
        if ($layoutPlaceholder) {
            $layoutPlaceholderType = $layoutPlaceholder.PlaceholderType.ToString()
            $boundsBox = Set-OfficePowerPointLayoutPlaceholder -Document $presentation -TemplateLayout $selectedLayout -Type $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -X 48 -Y 36 -Width 620 -Height 180 -PassThru
            $boundsBox.Left.ToPoints() | Should -Be 48
            $boundsBox.Top.ToPoints() | Should -Be 36
            $boundsBox.Width.ToPoints() | Should -Be 620
            $boundsBox.Height.ToPoints() | Should -Be 180

            $marginsBox = Set-OfficePowerPointLayoutPlaceholder -Document $presentation -TemplateLayout $selectedLayout -Type $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -TextMargin 8,12,8,12 -PassThru
            $marginsBox.TextMarginLeft.ToPoints() | Should -Be 12
            $marginsBox.TextMarginTop.ToPoints() | Should -Be 8
            $marginsBox.TextMarginRight.ToPoints() | Should -Be 12
            $marginsBox.TextMarginBottom.ToPoints() | Should -Be 8

            $styleBox = Set-OfficePowerPointLayoutPlaceholder -Document $presentation -TemplateLayout $selectedLayout -Type $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Style Body -Size 18 -Bold -PassThru
            $styleBox.FontSize.ToPoints() | Should -Be 18
            $styleBox.Bold | Should -BeTrue
        }

        Save-OfficePowerPoint -Document $presentation
        $presentation.Dispose()

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $reloadedSlide = Get-OfficePowerPointSlide -Document $reloaded -Slide 1
            $reloadedLayoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $reloadedSlide
            if ($layoutPlaceholder) {
                $reloadedLayoutPlaceholders.Count | Should -BeGreaterThan 0
                $reloadedLayoutPlaceholder = Get-OfficePowerPointLayoutPlaceholder -Slide $reloadedSlide -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex
                $reloadedLayoutPlaceholder | Should -Not -BeNullOrEmpty
                $reloadedLayoutPlaceholder.Bounds | Should -Not -BeNullOrEmpty
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.Left.ToPoints() - 48) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.Top.ToPoints() - 36) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.Width.ToPoints() - 620) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.Height.ToPoints() - 180) | Should -BeLessThan 0.1
            }
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'removes slides and preserves the remaining title' {
        $path = Join-Path $TestDrive 'PowerPointRemoval.pptx'
        $presentation = New-OfficePowerPoint -Path $path

        $firstSlide = $presentation.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $firstSlide -Title 'Remove me'

        $secondSlide = $presentation.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $secondSlide -Title 'Keep me'
        Set-OfficePowerPointPlaceholderText -Slide $secondSlide -PlaceholderType Title -Text 'Keep me v2'

        Save-OfficePowerPoint -Document $presentation
        $presentation.Dispose()

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            Remove-OfficePowerPointSlide -Document $reloaded -Slide 1 -Confirm:$false
            Save-OfficePowerPoint -Document $reloaded
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }

        $afterRemoval = Open-OfficePowerPoint -Path $path
        try {
            $afterRemoval.Slides.Count | Should -Be 1
            $remainingSlide = Get-OfficePowerPointSlide -Document $afterRemoval -Slide 1
            $remainingPlaceholder = Get-OfficePowerPointPlaceholder -Slide $remainingSlide -PlaceholderType Title
            if (-not $remainingPlaceholder) {
                $remainingPlaceholder = Get-OfficePowerPointPlaceholder -Slide $remainingSlide -PlaceholderType CenteredTitle
            }
            $remainingPlaceholder.Text | Should -Be 'Keep me v2'
        } finally {
            if ($afterRemoval) {
                $afterRemoval.Dispose()
            }
        }
    }

    It 'copies a slide and preserves its content' {
        $path = Join-Path $TestDrive 'PowerPointCopy.pptx'
        $presentation = New-OfficePowerPoint -Path $path

        $slide1 = $presentation.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $slide1 -Title 'Quarterly Overview'
        Add-OfficePowerPointTextBox -Slide $slide1 -Text 'Revenue and margin summary' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $slide1 -Text 'Reuse this for the board deck.'

        $slide2 = $presentation.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $slide2 -Title 'Closing Slide'

        $copiedSlide = Copy-OfficePowerPointSlide -Document $presentation -Slide 1 -Position 2 -PassThru
        $copiedSlide | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Document $presentation
        $presentation.Dispose()

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -Be 3

            $copied = Get-OfficePowerPointSlide -Document $reloaded -Slide 2
            $copiedTitle = Get-OfficePowerPointPlaceholder -Slide $copied -PlaceholderType Title
            if (-not $copiedTitle) {
                $copiedTitle = Get-OfficePowerPointPlaceholder -Slide $copied -PlaceholderType CenteredTitle
            }

            $copiedTitle.Text | Should -Be 'Quarterly Overview'
            (Get-OfficePowerPointNotes -Slide $copied).Text | Should -Be 'Reuse this for the board deck.'
            ($copied.Shapes.Count -gt 0) | Should -BeTrue

            $lastSlide = Get-OfficePowerPointSlide -Document $reloaded -Slide 3
            $lastTitle = Get-OfficePowerPointPlaceholder -Slide $lastSlide -PlaceholderType Title
            if (-not $lastTitle) {
                $lastTitle = Get-OfficePowerPointPlaceholder -Slide $lastSlide -PlaceholderType CenteredTitle
            }

            $lastTitle.Text | Should -Be 'Closing Slide'
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'supports sections, text replacement, and slide import' {
        $sourcePath = Join-Path $TestDrive 'PowerPointSourceDeck.pptx'
        $targetPath = Join-Path $TestDrive 'PowerPointTargetDeck.pptx'

        $source = New-OfficePowerPoint -Path $sourcePath
        $sourceSlide = $source.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $sourceSlide -Title 'FY24 Imported'
        Add-OfficePowerPointTextBox -Slide $sourceSlide -Text 'FY24 details from source' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $sourceSlide -Text 'FY24 source notes'
        Save-OfficePowerPoint -Document $source

        $target = New-OfficePowerPoint -Path $targetPath
        $introSlide = $target.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $introSlide -Title 'FY24 Overview'
        Add-OfficePowerPointTextBox -Slide $introSlide -Text 'FY24 summary for leadership' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $introSlide -Text 'FY24 note for intro'

        $resultsSlide = $target.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $resultsSlide -Title 'FY24 Results'
        Add-OfficePowerPointTextBox -Slide $resultsSlide -Text 'FY24 results body' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null

        $introSection = Add-OfficePowerPointSection -Document $target -Name 'Intro' -Slide 1 -PassThru
        $introSection.Name | Should -Be 'Intro'
        $resultsSection = Add-OfficePowerPointSection -Document $target -Name 'Results' -Slide 2 -PassThru
        $resultsSection.Name | Should -Be 'Results'

        $renamedSection = Rename-OfficePowerPointSection -Document $target -Name 'Results' -NewName 'Deep Dive' -PassThru
        $renamedSection.Name | Should -Be 'Deep Dive'

        $sections = @(Get-OfficePowerPointSection -Document $target)
        $sections.Count | Should -Be 2
        ($sections | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 0
        ($sections | Where-Object Name -eq 'Deep Dive').SlideIndices | Should -Contain 1

        $replacements = Update-OfficePowerPointText -Presentation $target -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes -PassThru
        $replacements | Should -BeGreaterThan 0

        $importedSlide = Import-OfficePowerPointSlide -Presentation $target -SourcePath $sourcePath -SourceIndex 0 -InsertAt 1
        $importedSlide | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Document $target

        $reloaded = Open-OfficePowerPoint -Path $targetPath
        try {
            $reloaded.Slides.Count | Should -Be 3

            $slide0 = Get-OfficePowerPointSlide -Document $reloaded -Slide 1
            $title0 = Get-OfficePowerPointPlaceholder -Slide $slide0 -PlaceholderType Title
            if (-not $title0) {
                $title0 = Get-OfficePowerPointPlaceholder -Slide $slide0 -PlaceholderType CenteredTitle
            }
            $title0.Text | Should -Be 'FY25 Overview'
            (Get-OfficePowerPointNotes -Slide $slide0).Text | Should -Be 'FY25 note for intro'

            $slide1 = Get-OfficePowerPointSlide -Document $reloaded -Slide 2
            $title1 = Get-OfficePowerPointPlaceholder -Slide $slide1 -PlaceholderType Title
            if (-not $title1) {
                $title1 = Get-OfficePowerPointPlaceholder -Slide $slide1 -PlaceholderType CenteredTitle
            }
            $title1.Text | Should -Be 'FY24 Imported'
            (Get-OfficePowerPointNotes -Slide $slide1).Text | Should -Be 'FY24 source notes'

            $sectionInfo = @(Get-OfficePowerPointSection -Document $reloaded)
            $sectionInfo.Count | Should -Be 2
            ($sectionInfo | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 0
            ($sectionInfo | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 1
            ($sectionInfo | Where-Object Name -eq 'Deep Dive').SlideIndices | Should -Contain 2

            $updatedReplacements = Update-OfficePowerPointText -Presentation $reloaded -OldValue 'FY24' -NewValue 'FY26' -PassThru
            $updatedReplacements | Should -BeGreaterThan 0
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }
}
