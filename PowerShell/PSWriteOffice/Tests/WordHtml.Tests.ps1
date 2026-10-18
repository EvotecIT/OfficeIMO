BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Word HTML conversions' {
    It 'converts Word to HTML string' {
        $docPath = Join-Path $TestDrive 'HtmlSource.docx'
        New-OfficeWord -Path $docPath {
            Add-OfficeWordSection {
                Add-OfficeWordParagraph -Text 'Hello HTML'
            }
        } | Out-Null

        $html = ConvertTo-OfficeWordHtml -Path $docPath
        $html | Should -Match 'Hello HTML'
    }

    It 'converts HTML to Word document' {
        $html = '<h1>Hello HTML</h1><p>Roundtrip</p>'
        $docPath = Join-Path $TestDrive 'HtmlRoundtrip.docx'

        ConvertFrom-OfficeWordHtml -Html $html -OutputPath $docPath | Out-Null
        Test-Path $docPath | Should -BeTrue

        $document = Get-OfficeWord -Path $docPath -ReadOnly
        try {
            $document.Paragraphs.Count | Should -BeGreaterThan 0
        } finally {
            $document.Dispose()
        }
    }
}
