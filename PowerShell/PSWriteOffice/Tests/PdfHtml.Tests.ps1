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

        ConvertFrom-OfficePdfHtml -Html '<h1>HTML Report</h1><p>Ready for archive.</p>' -OutputPath $path -PassThru |
            Should -BeOfType System.IO.FileInfo

        Test-Path $path | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'HTML Report'
        $text | Should -Match 'Ready for archive'
    }

    It 'accumulates piped HTML lines into one PDF document' {
        $path = Join-Path $TestDrive 'pipeline-html-report.pdf'

        @(
            '<h1>Pipeline HTML</h1>'
            '<p>Full document body.</p>'
        ) | ConvertFrom-OfficePdfHtml -OutputPath $path | Out-Null

        Test-Path $path | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'Pipeline HTML'
        $text | Should -Match 'Full document body'
    }

    It 'converts an HTML file when InputPath is named' {
        $htmlPath = Join-Path $TestDrive 'source.html'
        $path = Join-Path $TestDrive 'html-file-report.pdf'
        Set-Content -Path $htmlPath -Value '<h1>File HTML</h1><p>Loaded from disk.</p>' -Encoding UTF8

        ConvertFrom-OfficePdfHtml -Path $htmlPath -OutputPath $path -PassThru |
            Should -BeOfType System.IO.FileInfo

        Test-Path $path | Should -BeTrue
        $text = Get-OfficePdfText -Path $path
        $text | Should -Match 'File HTML'
        $text | Should -Match 'Loaded from disk'
    }

    It 'converts a PDF file to semantic HTML' {
        $path = Join-Path $TestDrive 'source.pdf'
        $htmlPath = Join-Path $TestDrive 'source.html'
        ConvertFrom-OfficePdfHtml -Html '<h1>PDF To HTML</h1><p>Logical export.</p>' -OutputPath $path | Out-Null

        ConvertTo-OfficePdfHtml -Path $path -OutputPath $htmlPath |
            Should -BeOfType System.IO.FileInfo

        $html = Get-Content -Path $htmlPath -Raw
        $html | Should -Match '<html'
        $html | Should -Match 'PDF To HTML'
        $html | Should -Match 'Logical export'
    }
}
