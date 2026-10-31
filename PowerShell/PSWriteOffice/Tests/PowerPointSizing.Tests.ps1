BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'PowerPoint slide sizing through the module' {
    It 'accepts physical lengths and preserves the dimensions and transition after saving' {
        $path = Join-Path $TestDrive 'Custom-size.pptx'
        New-OfficePowerPoint -Path $path {
            $size = Set-OfficePowerPointSlideSize -Width '25.4cm' -Height '14cm' -PassThru
            [math]::Round($size.Width.ToCentimeters(), 1) | Should -Be 25.4
            [math]::Round($size.Height.ToCentimeters(), 1) | Should -Be 14.0
            $slide = PowerPointSlide -Title 'Transition demo' -PassThru
            Set-OfficePowerPointSlide -Slide $slide -Transition Fade
        }
        $document = Open-OfficePowerPoint -LiteralPath $path
        try {
            [math]::Round($document.SlideSize.Width.ToCentimeters(), 1) | Should -Be 25.4
            [math]::Round($document.SlideSize.Height.ToCentimeters(), 1) | Should -Be 14.0
            $document.Slides[0].Transition | Should -Be 'Fade'
            @($document.ValidateDocument()).Count | Should -Be 0
        } finally {
            Close-OfficePowerPoint -Document $document
        }
    }

    It 'runs the widescreen transitions example using the current commands' {
        $file = & (Join-Path $PSScriptRoot '..\Examples\PowerPoint\Example-PowerPointTransitionsAndSizing.ps1') -OutputDirectory $TestDrive
        $file | Should -BeOfType ([System.IO.FileInfo])
        $document = Open-OfficePowerPoint -LiteralPath $file.FullName
        try {
            $document.Slides.Count | Should -Be 2
            [math]::Round($document.SlideSize.AspectRatio, 2) | Should -Be 1.78
            $document.Slides[0].Transition | Should -Be 'Fade'
            $document.Slides[1].Transition | Should -Be 'Morph'
            @($document.ValidateDocument()).Count | Should -Be 0
        } finally {
            Close-OfficePowerPoint -Document $document
        }
    }
}
