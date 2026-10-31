BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'PowerPoint HTML output controls' {
    It 'preserves slide content while omitting requested review text for <Name>' -ForEach @(
        @{ Name = 'semantic slides'; Profile = 'SemanticSlides' }
        @{ Name = 'visual review'; Profile = 'VisualReview' }
    ) {
        $deckPath = Join-Path $TestDrive ($Profile + '.pptx')
        $htmlPath = Join-Path $TestDrive ($Profile + '.html')
        $presentation = New-OfficePowerPoint -Path $deckPath -ErrorAction Stop
        try {
            $presentation.SlideSize.SetSize([OfficeIMO.OfficeLength]::Points(160), [OfficeIMO.OfficeLength]::Points(100))
            $slide = $presentation.AddSlide()
            Add-OfficePowerPointTextBox -Slide $slide -Text 'Visible slide content' -X 10 -Y 10 -Width 140 -Height 50 -ErrorAction Stop
            $slide.AddRectangle([OfficeIMO.OfficeLength]::Points(500), [OfficeIMO.OfficeLength]::Points(500), [OfficeIMO.OfficeLength]::Points(20), [OfficeIMO.OfficeLength]::Points(20), 'Outside slide bounds') | Out-Null

            $defaults = ConvertTo-OfficePowerPointHtml -Presentation $presentation -Profile $Profile -ErrorAction Stop
            $defaults | Should -Match '<h1>PowerPoint'
            $defaults | Should -Match '<h2>Slide 1</h2>'
            $defaults | Should -Match 'class="officeimo-feature officeimo-masters"'
            $defaults | Should -Match '<summary>Extraction proof</summary>'
            if ($Profile -eq 'VisualReview') {
                $defaults | Should -Match '<ul class="officeimo-snapshot-diagnostics">'
                $defaults | Should -Match 'invalid-shape-bounds'
            }

            $outputChoices = @{
                Profile = $Profile
                NoHeading = $true
                NoSlideHeadings = $true
                NoRenderDiagnostics = $true
                NoMasterInventory = $true
                NoExtractionProof = $true
                ErrorAction = 'Stop'
            }
            $fromPresentation = ConvertTo-OfficePowerPointHtml -Presentation $presentation @outputChoices
            Save-OfficePowerPoint -Document $presentation -ErrorAction Stop
            Close-OfficePowerPoint -Document $presentation -ErrorAction Stop
            $presentation = $null

            $saved = ConvertTo-OfficePowerPointHtml -Path $deckPath -OutputPath $htmlPath -PassThru @outputChoices
            $saved | Should -BeOfType System.IO.FileInfo
            $saved.FullName | Should -Be $htmlPath
            foreach ($html in @($fromPresentation, (Get-Content -LiteralPath $htmlPath -Raw))) {
                $html | Should -Match 'Visible slide content'
                $html | Should -Match '<section class="officeimo-slide"'
                $html | Should -Match '<title>PowerPoint'
                $html | Should -Not -Match '<h1>'
                $html | Should -Not -Match '<h2>Slide [0-9]+</h2>'
                $html | Should -Not -Match '<ul class="officeimo-snapshot-diagnostics">'
                $html | Should -Not -Match 'class="officeimo-feature officeimo-masters"'
                $html | Should -Not -Match '<summary>Extraction proof</summary>'
                if ($Profile -eq 'VisualReview') {
                    $html | Should -Match 'class="officeimo-slide-canvas"'
                    $html | Should -Match '<svg'
                }
            }
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Document $presentation -ErrorAction SilentlyContinue
            }
        }
    }
}