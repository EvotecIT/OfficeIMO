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
            $layoutPlaceholders = @(Get-OfficePowerPointLayoutPlaceholder -Slide $slide)
            $layoutPlaceholders.Count | Should -BeGreaterThan 0
            $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
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
            $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
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

        Set-OfficePowerPointThemeColor -Presentation $presentation -Colors @{
            Accent1 = '#C00000'
            Accent2 = '00B0F0'
        } -AllMasters
        Set-OfficePowerPointThemeFonts -Presentation $presentation -MajorLatin 'Aptos' -MinorLatin 'Calibri' -AllMasters
        Set-OfficePowerPointThemeName -Presentation $presentation -Name 'Contoso Theme' -AllMasters

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

            $contentBox = Get-OfficePowerPointLayoutBox -Presentation $presentation -MarginCm 1.5
            [math]::Round($contentBox.Left.ToCentimeters(), 2) | Should -Be 1.5
            [math]::Round($contentBox.Top.ToCentimeters(), 2) | Should -Be 1.5
            [math]::Round($contentBox.Width.ToCentimeters(), 2) | Should -Be 27
            [math]::Round($contentBox.Height.ToCentimeters(), 2) | Should -Be 17

            $columns = @(Get-OfficePowerPointLayoutBox -Presentation $presentation -ColumnCount 2 -MarginCm 1.5 -GutterCm 1.0)
            $columns.Count | Should -Be 2
            [math]::Round($columns[0].Width.ToCentimeters(), 2) | Should -Be 13
            [math]::Round($columns[1].Left.ToCentimeters(), 2) | Should -Be 15.5

            $rows = @(Get-OfficePowerPointLayoutBox -Presentation $presentation -RowCount 2 -MarginCm 1.5 -GutterCm 0.5)
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

    It 'previews and renders OfficeIMO designer deck plans with native handles' {
        $path = Join-Path $TestDrive 'PowerPointDesignerDeck.pptx'
        $plan = New-OfficePowerPointDeckPlan
        $plan.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointDeckPlan'
        Add-OfficePowerPointPlanSection -Plan $plan -Title 'Reliability Showcase' -Subtitle 'From script to polished deck'
        Add-OfficePowerPointPlanProcess -Plan $plan -Title 'Delivery Flow' -Steps @(
            [PSCustomObject]@{ Number = '01'; Title = 'Author'; Body = 'Describe the story once.' }
            [PSCustomObject]@{ Number = '02'; Title = 'Render'; Body = 'Let OfficeIMO compose the layout.' }
            [PSCustomObject]@{ Number = '03'; Title = 'Validate'; Body = 'Inspect and test the package.' }
        )
        Add-OfficePowerPointPlanCardGrid -Plan $plan -Title 'Proof Points' -Cards @(
            [PSCustomObject]@{ Title = 'Useful'; Items = @('Semantic input', 'Readable output') }
            [PSCustomObject]@{ Title = 'Fast'; Items = @('One cmdlet bridge', 'Deterministic seed') }
        )

        $preview = @(Add-OfficePowerPointDesignerDeck -Plan $plan -AccentColor '#008C95' -Seed 'designer-test' -Purpose 'technical service brief' -Preview -PassThru)
        $preview.Count | Should -BeGreaterThan 0

        $presentation = New-OfficePowerPoint -Path $path
        try {
            $renderedSlides = @(Add-OfficePowerPointDesignerDeck -Presentation $presentation -Plan $plan -AccentColor '#008C95' -Seed 'designer-test' -Purpose 'technical service brief' -Name 'Designer Test' -LayoutStrategy ContentFirst -PassThru)
            $renderedSlides.Count | Should -BeGreaterThan 2
            foreach ($renderedSlide in $renderedSlides) {
                $renderedSlide.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.PowerPointSlide'
            }
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

    It 'rejects designer plan data arrays that contain only null values' {
        $mapper = [PSWriteOffice.Cmdlets.PowerPoint.AddOfficePowerPointPlanProcessCommand].Assembly.GetType('PSWriteOffice.Services.PowerPoint.PowerPointDesignerDataMapper')
        $method = $mapper.GetMethod('ToProcessSteps', [System.Reflection.BindingFlags] 'Public,Static')

        try {
            $method.Invoke($null, (, [object[]] @($null)))
            throw 'Expected mapper to reject all-null process data.'
        } catch {
            $_.Exception.InnerException.Message | Should -Be 'Process steps require at least one item.'
        }
    }

    It 'normalizes blank optional designer colors to null' {
        $mapper = [PSWriteOffice.Cmdlets.PowerPoint.AddOfficePowerPointPlanProcessCommand].Assembly.GetType('PSWriteOffice.Services.PowerPoint.PowerPointDesignerDataMapper')
        $method = $mapper.GetMethod('ToCards', [System.Reflection.BindingFlags] 'Public,Static')

        $cards = $method.Invoke($null, (, [object[]] @([PSCustomObject]@{ Title = 'Blank color'; Items = @('Still valid'); AccentColor = '   ' })))
        $cards[0].AccentColor | Should -BeNullOrEmpty
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
            $columnChart = Add-OfficePowerPointChart -Slide $slide1 -Data $rows -Category Month -Value Sales, Profit -Title 'Sales vs Profit' -X 40 -Y 120 -Width 360 -Height 220 -PassThru
            $columnChart | Should -Not -BeNullOrEmpty
            @($slide1.Charts).Count | Should -Be 1

            $slide2 = $presentation.AddSlide(0, 1)
            Set-OfficePowerPointSlide -Slide $slide2 -Title 'Pie Chart' | Out-Null
            $pieChart = Add-OfficePowerPointChart -Slide $slide2 -Type Pie -InputObject $rows -Category Month -Value Sales -Title 'Sales Mix' -X 40 -Y 120 -Width 320 -Height 220 -PassThru
            $pieChart | Should -Not -BeNullOrEmpty
            @($slide2.Charts).Count | Should -Be 1

            $slide3 = $presentation.AddSlide(0, 1)
            Set-OfficePowerPointSlide -Slide $slide3 -Title 'Scatter Chart' | Out-Null
            $scatterChart = Add-OfficePowerPointChart -Slide $slide3 -Type Scatter -Data $rows -XProperty MonthNumber -YProperty Sales, Profit -Title 'Trend Scatter' -X 40 -Y 120 -Width 360 -Height 220 -PassThru
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
            (Get-OfficePowerPointShape -Presentation $reloaded | Where-Object Kind -eq 'Chart') | Should -HaveCount 3
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }
}
