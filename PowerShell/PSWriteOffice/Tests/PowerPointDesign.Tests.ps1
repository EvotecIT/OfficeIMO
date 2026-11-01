BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'PowerPointDesign workflows' {
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

    It 'creates a presentation through content blocks and edits native slide and layout handles' {
        $path = Join-Path $TestDrive 'PowerPointNativeLayout.pptx'
        New-OfficePowerPoint -Path $path {
            $layout = Get-OfficePowerPointLayout | Where-Object Type -eq Text | Select-Object -First 1
            $layout | Should -Not -BeNullOrEmpty
            $slide = Add-OfficePowerPointSlide -Title 'Native Slide' -TemplateLayout $layout -PassThru {
                Add-OfficePowerPointParagraph -Text 'Hello from content blocks'
            }
            $slide.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointSlide'
            $textBox = Add-OfficePowerPointTextBox -Slide $slide -Text 'Hello from native handles' -X 80 -Y 150 -Width 240 -Height 60 -PassThru
            $textBox.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointTextBox'
            $textBox.Text | Should -Be 'Hello from native handles'
            $layoutPlaceholders = @(Get-OfficePowerPointPlaceholder -FromLayout -Slide $slide)
            $layoutPlaceholders.Count | Should -BeGreaterThan 0
            $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -Type Title
            $placeholder.Text | Should -Be 'Native Slide'
        }

        $presentation = Open-OfficePowerPoint -Path $path
        try {
            $slide = Get-OfficePowerPointSlide -Document $presentation -Slide 1
            $layout = Get-OfficePowerPointLayout -Document $presentation |
                Where-Object LayoutIndex -eq $slide.LayoutIndex | Select-Object -First 1
            $layout | Should -Not -BeNullOrEmpty
            Set-OfficePowerPointLayoutPlaceholder -Document $presentation -TemplateLayout $layout -Type Title `
                -X 60 -Y 40 -Width 600 -Height 120 -CreateIfMissing
            Set-OfficePowerPointLayoutPlaceholder -Document $presentation -TemplateLayout $layout -Type Title `
                -TextMargin 6,8,6,8 -CreateIfMissing
            Set-OfficePowerPointLayoutPlaceholder -Document $presentation -TemplateLayout $layout -Type Title `
                -Style Title -Size 32 -Bold -CreateIfMissing

            Save-OfficePowerPoint -Document $presentation
        } finally {
            Close-OfficePowerPoint -Document $presentation
        }

        Test-Path $path | Should -BeTrue
        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $slide = Get-OfficePowerPointSlide -Document $reloaded -Slide 1
            $summary = Get-OfficePowerPointSlide -Document $reloaded -Slide 1 -Summary
            $summary.LayoutPlaceholderCount | Should -BeGreaterThan 0
            $summary.Title | Should -Be 'Native Slide'
            $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -Type Title
            $placeholder.Text | Should -Be 'Native Slide'
            $slide.TextBoxes.Text | Should -Contain 'Hello from native handles'
            $slide.TextBoxes.Text | Should -Contain 'Hello from content blocks'
        } finally {
            Close-OfficePowerPoint -Document $reloaded
        }
    }

    It 'supports theme inspection, theme updates, and slide layout switching' {
        $path = Join-Path $TestDrive 'PowerPointThemeAndLayout.pptx'
        $presentation = New-OfficePowerPoint -Path $path

        $slide = $presentation.AddSlide(0, 1)
        Set-OfficePowerPointSlide -Slide $slide -Title 'Theme Demo' | Out-Null

        $layouts = Get-OfficePowerPointLayout -Document $presentation
        $layouts.Count | Should -BeGreaterThan 1
        $currentLayoutIndex = $slide.LayoutIndex
        $alternativeLayout = $layouts | Where-Object { $_.LayoutIndex -ne $currentLayoutIndex } | Select-Object -First 1
        $alternativeLayout | Should -Not -BeNullOrEmpty

        $themeBefore = Get-OfficePowerPointTheme -Document $presentation
        $themeBefore.MasterIndex | Should -Be 0
        $themeBefore.Colors.Count | Should -BeGreaterThan 0

        Set-OfficePowerPointTheme -Document $presentation -Colors @{
            Accent1 = '#C00000'
            Accent2 = '00B0F0'
        } -HeadingFont 'Aptos' -BodyFont 'Calibri' -Name 'Contoso Theme' -AllMasters

        $slide | Set-OfficePowerPointSlide -Document $presentation -TemplateLayout $alternativeLayout | Out-Null

        $themeAfter = Get-OfficePowerPointTheme -Document $presentation
        $themeAfter.ThemeName | Should -Be 'Contoso Theme'
        $accent1Color = Get-TestPSWriteOfficeEnumValue -AssemblyName 'OfficeIMO.PowerPoint' -TypeName 'OfficeIMO.PowerPoint.PowerPointThemeColor' -Name 'Accent1' -CommandName 'New-OfficePowerPoint'
        $accent2Color = Get-TestPSWriteOfficeEnumValue -AssemblyName 'OfficeIMO.PowerPoint' -TypeName 'OfficeIMO.PowerPoint.PowerPointThemeColor' -Name 'Accent2' -CommandName 'New-OfficePowerPoint'
        $themeAfter.Colors[$accent1Color] | Should -Be 'C00000'
        $themeAfter.Colors[$accent2Color] | Should -Be '00B0F0'
        $themeAfter.MajorLatin | Should -Be 'Aptos'
        $themeAfter.MinorLatin | Should -Be 'Calibri'
        $slide.LayoutIndex | Should -Be $alternativeLayout.LayoutIndex

        Save-OfficePowerPoint -Document $presentation
        $presentation.Dispose()

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $reloadedTheme = Get-OfficePowerPointTheme -Document $reloaded
            $reloadedTheme.ThemeName | Should -Be 'Contoso Theme'
            $reloadedTheme.Colors[$accent1Color] | Should -Be 'C00000'
            $reloadedTheme.Colors[$accent2Color] | Should -Be '00B0F0'
            $reloadedTheme.MajorLatin | Should -Be 'Aptos'
            $reloadedTheme.MinorLatin | Should -Be 'Calibri'

            $reloadedSlide = Get-OfficePowerPointSlide -Document $reloaded -Slide 1
            $reloadedSlide.LayoutIndex | Should -Be $alternativeLayout.LayoutIndex
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'supports slide backgrounds and layout box helpers' {
        $path = Join-Path $TestDrive 'PowerPointBackgroundsAndLayout.pptx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive
        $presentation = $null

        try {
            $presentation = New-OfficePowerPoint -Path $path

            Set-OfficePowerPointSlideSize -Document $presentation -Width '30cm' -Height '20cm' | Out-Null

            $colorSlide = $presentation.AddSlide(0, 1)
            Set-OfficePowerPointSlide -Slide $colorSlide -Title 'Layout Demo' | Out-Null
            $updatedSlide = Set-OfficePowerPointSlide -Slide $colorSlide -Background '#F4F7FB' -PassThru
            $updatedSlide | Should -Be $colorSlide
            $updatedSlide.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointSlide'
            $colorSlide.BackgroundColor | Should -Be 'f4f7fb'

            $contentBox = Get-OfficePowerPointLayoutBox -Document $presentation -Margin '1.5cm'
            [math]::Round($contentBox.Left.ToCentimeters(), 2) | Should -Be 1.5
            [math]::Round($contentBox.Top.ToCentimeters(), 2) | Should -Be 1.5
            [math]::Round($contentBox.Width.ToCentimeters(), 2) | Should -Be 27
            [math]::Round($contentBox.Height.ToCentimeters(), 2) | Should -Be 17

            $columns = @(Get-OfficePowerPointLayoutBox -Document $presentation -Columns 2 -Margin '1.5cm' -Gutter '1cm')
            $columns.Count | Should -Be 2
            [math]::Round($columns[0].Width.ToCentimeters(), 2) | Should -Be 13
            [math]::Round($columns[1].Left.ToCentimeters(), 2) | Should -Be 15.5

            $rows = @(Get-OfficePowerPointLayoutBox -Document $presentation -Rows 2 -Margin '1.5cm' -Gutter '0.5cm')
            $rows.Count | Should -Be 2
            [math]::Round($rows[0].Height.ToCentimeters(), 2) | Should -Be 8.25
            [math]::Round($rows[1].Top.ToCentimeters(), 2) | Should -Be 10.25

            Add-OfficePowerPointTextBox -Slide $colorSlide -Text 'Left column' -X $columns[0].Left.ToPoints() -Y $columns[0].Top.ToPoints() -Width $columns[0].Width.ToPoints() -Height 40 | Out-Null
            Add-OfficePowerPointTextBox -Slide $colorSlide -Text 'Right column' -X $columns[1].Left.ToPoints() -Y $columns[1].Top.ToPoints() -Width $columns[1].Width.ToPoints() -Height 40 | Out-Null

            $imageSlide = $presentation.AddSlide(0, 1)
            Set-OfficePowerPointSlide -Slide $imageSlide -Title 'Image Background' | Out-Null
            Set-OfficePowerPointSlide -Slide $imageSlide -BackgroundImage $imagePath | Out-Null
            $imageSlide.BackgroundColor | Should -BeNullOrEmpty

            Save-OfficePowerPoint -Document $presentation
        } finally {
            if ($presentation) {
                $presentation.Dispose()
            }
        }

        $colorSlideXml = Get-ZipXmlDocumentLocal -Path $path -Entry (Get-TestPowerPointSlideEntry -Path $path -Index 0)
        $colorSlideXml.OuterXml | Should -Match 'rgbClr val="f4f7fb"'

        $imageSlideXml = Get-ZipXmlDocumentLocal -Path $path -Entry (Get-TestPowerPointSlideEntry -Path $path -Index 1)
        $imageSlideXml.OuterXml | Should -Match '<a:blip '
        $imageSlideXml.OuterXml | Should -Match 'r:embed='

        $entries = @(Get-ZipEntriesLocal -Path $path)
        ($entries | Where-Object { $_ -like 'ppt/media/*' }).Count | Should -BeGreaterThan 0

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $reloadedColorSlide = Get-OfficePowerPointSlide -Document $reloaded -Slide 1
            $reloadedColorSlide.BackgroundColor | Should -Be 'f4f7fb'

            $reloadedImageSlide = Get-OfficePowerPointSlide -Document $reloaded -Slide 2
            $reloadedImageSlide.BackgroundColor | Should -BeNullOrEmpty
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'composes designed slides directly and reports their native layout summaries' {
        $path = Join-Path $TestDrive 'PowerPointDesignerDeck.pptx'
        $presentation = New-OfficePowerPoint -Path $path -Theme Modern -Accent '#008C95' -Seed 'designer-test' `
            -Purpose 'technical service brief' -Style TechnicalMap -LayoutStrategy ContentFirst
        try {
            $renderedSlides = @(
                PowerPointSlide -Document $presentation -Title 'Reliability Showcase' -Subtitle 'From script to polished deck' -Layout Title -PassThru
                PowerPointSlide -Document $presentation -Title 'Delivery Flow' -PassThru {
                    PowerPointStep 'Author' 'Describe the story once.'
                    PowerPointStep 'Render' 'Compose the layout.'
                    PowerPointStep 'Validate' 'Inspect and test the package.'
                }
                PowerPointSlide -Document $presentation -Title 'Proof Points' -PassThru {
                    PowerPointCard 'Useful' -Items 'Semantic input', 'Readable output'
                    PowerPointCard 'Fast' -Items 'One content block', 'Deterministic seed'
                }
            )
            $renderedSlides.Count | Should -BeGreaterThan 2
            foreach ($renderedSlide in $renderedSlides) {
                $renderedSlide.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointSlide'
            }
            $summary = @(Get-OfficePowerPointSlide -Document $presentation -Summary)
            $summary.Count | Should -Be $renderedSlides.Count
            ($summary | Where-Object { $_.ShapeCount -gt 0 }).Count | Should -Be $renderedSlides.Count
            Save-OfficePowerPoint -Document $presentation
        } finally {
            Close-OfficePowerPoint -Document $presentation
        }

        $entries = @(Get-ZipEntriesLocal -Path $path)
        ($entries | Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' }).Count | Should -BeGreaterThan 2

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -BeGreaterThan 2
            $summary = @(Get-OfficePowerPointSlide -Document $reloaded -Summary)
            $summary.Count | Should -BeGreaterThan 2
            ($summary | Where-Object { $_.ShapeCount -gt 0 }).Count | Should -BeGreaterThan 2
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'supports PowerPoint charts from object data' {
        $path = Join-Path $TestDrive 'PowerPointCharts.pptx'
        $presentation = $null
        $rows = @(
            [PSCustomObject]@{ Month = 'Jan'; MonthNumber = 1; Sales = 10; Profit = 4 }
            [PSCustomObject]@{ Month = 'Feb'; MonthNumber = 2; Sales = 14; Profit = 6 }
            [PSCustomObject]@{ Month = 'Mar'; MonthNumber = 3; Sales = 18; Profit = 8 }
        )

        try {
            $presentation = New-OfficePowerPoint -Path $path

            $slide1 = $presentation.AddSlide(0, 1)
            Set-OfficePowerPointSlide -Slide $slide1 -Title 'Column Chart' | Out-Null
            $columnChart = Add-OfficePowerPointChart -Slide $slide1 -InputObject $rows -Category Month -Value Sales, Profit -Title 'Sales vs Profit' -X 40 -Y 120 -Width 360 -Height 220 -PassThru
            $columnChart | Should -Not -BeNullOrEmpty
            @($slide1.Charts).Count | Should -Be 1

            $slide2 = $presentation.AddSlide(0, 1)
            Set-OfficePowerPointSlide -Slide $slide2 -Title 'Pie Chart' | Out-Null
            $pieChart = Add-OfficePowerPointChart -Slide $slide2 -Type Pie -InputObject $rows -Category Month -Value Sales -Title 'Sales Mix' -X 40 -Y 120 -Width 320 -Height 220 -PassThru
            $pieChart | Should -Not -BeNullOrEmpty
            @($slide2.Charts).Count | Should -Be 1

            $slide3 = $presentation.AddSlide(0, 1)
            Set-OfficePowerPointSlide -Slide $slide3 -Title 'Scatter Chart' | Out-Null
            $scatterChart = Add-OfficePowerPointChart -Slide $slide3 -Type Scatter -InputObject $rows -Category MonthNumber -Value Sales, Profit -Title 'Trend Scatter' -X 40 -Y 120 -Width 360 -Height 220 -PassThru
            $scatterChart | Should -Not -BeNullOrEmpty
            @($slide3.Charts).Count | Should -Be 1

            (Get-OfficePowerPointShape -Slide $slide1 | Where-Object Kind -eq 'Chart') | Should -HaveCount 1
            (Get-OfficePowerPointShape -Slide $slide2 | Where-Object Kind -eq 'Chart') | Should -HaveCount 1
            (Get-OfficePowerPointShape -Slide $slide3 | Where-Object Kind -eq 'Chart') | Should -HaveCount 1

            Save-OfficePowerPoint -Document $presentation
        } finally {
            if ($presentation) {
                $presentation.Dispose()
            }
        }

        $chartEntries = @(Get-ZipEntriesLocal -Path $path | Where-Object { $_ -match '^ppt/charts/chart\d+\.xml$' } | Sort-Object)
        $chartEntries.Count | Should -Be 3

        $chart1Xml = Get-ZipXmlDocumentLocal -Path $path -Entry $chartEntries[0]
        $chart1Xml.OuterXml | Should -Match '<c:barChart'
        $chart1Xml.OuterXml | Should -Match 'Sales vs Profit'

        $chart2Xml = Get-ZipXmlDocumentLocal -Path $path -Entry $chartEntries[1]
        $chart2Xml.OuterXml | Should -Match '<c:pieChart'
        $chart2Xml.OuterXml | Should -Match 'Sales Mix'

        $chart3Xml = Get-ZipXmlDocumentLocal -Path $path -Entry $chartEntries[2]
        $chart3Xml.OuterXml | Should -Match '<c:scatterChart'
        $chart3Xml.OuterXml | Should -Match 'Trend Scatter'

        $reloaded = Open-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -Be 3
            (Get-OfficePowerPointShape -Document $reloaded | Where-Object Kind -eq 'Chart') | Should -HaveCount 3
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }
}
