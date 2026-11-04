BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'PDF HTML cmdlets' {
    It 'converts HTML content to a readable PDF file' {
        $path = Join-Path $TestDrive 'html-report.pdf'

        Convert-OfficeDocument -To Pdf -From Html -Text '<h1>HTML Report</h1><p>Ready for archive.</p>' -OutputPath $path |
            Should -BeOfType OfficeIMO.OfficeConversionReport

        Test-Path $path | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'HTML Report'
        $text | Should -Match 'Ready for archive'
    }

    It 'accumulates piped HTML lines into one PDF document' {
        $path = Join-Path $TestDrive 'pipeline-html-report.pdf'

        @(
            '<h1>Pipeline HTML</h1>'
            '<p>Full document body.</p>'
        ) | Convert-OfficeDocument -To Pdf -From Html -Text { $_ } -OutputPath $path | Out-Null

        Test-Path $path | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'Pipeline HTML'
        $text | Should -Match 'Full document body'
    }

    It 'converts an HTML file when InputPath is named' {
        $htmlPath = Join-Path $TestDrive 'source.html'
        $path = Join-Path $TestDrive 'html-file-report.pdf'
        Set-Content -Path $htmlPath -Value '<h1>File HTML</h1><p>Loaded from disk.</p>' -Encoding UTF8

        Convert-OfficeDocument -To Pdf -Path $htmlPath -OutputPath $path |
            Should -BeOfType OfficeIMO.OfficeConversionReport

        Test-Path $path | Should -BeTrue
        $text = Read-OfficeDocument -As Text -Path $path
        $text | Should -Match 'File HTML'
        $text | Should -Match 'Loaded from disk'
    }

    It 'converts a PDF file to semantic HTML' {
        $path = Join-Path $TestDrive 'source.pdf'
        $htmlPath = Join-Path $TestDrive 'source.html'
        Convert-OfficeDocument -To Pdf -From Html -Text '<h1>PDF To HTML</h1><p>Logical export.</p>' -OutputPath $path | Out-Null

        Convert-OfficeDocument -To Html -Path $path -OutputPath $htmlPath |
            Should -BeOfType OfficeIMO.OfficeConversionReport

        $html = Get-Content -Path $htmlPath -Raw
        $html | Should -Match '<html'
        $html | Should -Match 'PDF To HTML'
        $html | Should -Match 'Logical export'
    }
}
