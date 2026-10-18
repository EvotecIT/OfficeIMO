BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'RTF cmdlets' {
    It 'converts HTML to RTF and RTF back to HTML' {
        $rtfPath = Join-Path $TestDrive 'HtmlSource.rtf'
        $htmlPath = Join-Path $TestDrive 'HtmlRoundtrip.html'

        ConvertTo-OfficeRtf -Html '<h1>Status</h1><p>Ready for HTML bridge</p>' -OutputPath $rtfPath -PassThru |
            Should -BeOfType System.IO.FileInfo

        Test-Path -LiteralPath $rtfPath | Should -BeTrue
        $html = ConvertFrom-OfficeRtf -Path $rtfPath -As Html
        $html | Should -Match 'Ready for HTML bridge'

        ConvertFrom-OfficeRtf -Path $rtfPath -As Html -OutputPath $htmlPath -PassThru |
            Should -BeOfType System.IO.FileInfo
        (Get-Content -LiteralPath $htmlPath -Raw) | Should -Match 'Ready for HTML bridge'
    }

    It 'converts Word to RTF and RTF back to Word' {
        $docPath = Join-Path $TestDrive 'WordSource.docx'
        $rtfPath = Join-Path $TestDrive 'WordSource.rtf'
        $roundtripPath = Join-Path $TestDrive 'WordRoundtrip.docx'

        New-OfficeWord -Path $docPath {
            Add-OfficeWordSection {
                Add-OfficeWordParagraph -Text 'Word RTF bridge'
            }
        } | Out-Null

        ConvertTo-OfficeRtf -WordPath $docPath -OutputPath $rtfPath -PassThru |
            Should -BeOfType System.IO.FileInfo
        ConvertFrom-OfficeRtf -Path $rtfPath -As Word -OutputPath $roundtripPath -PassThru |
            Should -BeOfType System.IO.FileInfo

        $document = Get-OfficeWord -Path $roundtripPath -ReadOnly
        try {
            ($document.Paragraphs | ForEach-Object Text) -join "`n" | Should -Match 'Word RTF bridge'
        } finally {
            $document.Dispose()
        }
    }

    It 'converts RTF to PDF and PDF back to RTF files' {
        $rtfPath = Join-Path $TestDrive 'PdfSource.rtf'
        $pdfPath = Join-Path $TestDrive 'PdfSource.pdf'
        $roundtripRtf = Join-Path $TestDrive 'PdfRoundtrip.rtf'

        New-OfficeRtf -Path $rtfPath -Text 'RTF PDF bridge' | Out-Null

        ConvertFrom-OfficeRtf -Path $rtfPath -As Pdf -OutputPath $pdfPath -PassThru |
            Should -BeOfType System.IO.FileInfo
        Test-Path -LiteralPath $pdfPath | Should -BeTrue
        ([System.IO.File]::ReadAllBytes($pdfPath)[0..3] -join ',') | Should -Be '37,80,68,70'

        ConvertTo-OfficeRtf -PdfPath $pdfPath -OutputPath $roundtripRtf -PassThru |
            Should -BeOfType System.IO.FileInfo
        (Get-OfficeRtf -Path $roundtripRtf).ToRtfLossless() | Should -Match 'RTF PDF bridge'
    }

    It 'creates and reads RTF files through OfficeIMO.Rtf' {
        $path = Join-Path $TestDrive 'Report.rtf'

        $file = New-OfficeRtf -Path $path -Text 'Summary', 'Ready for review' -PassThru

        $file | Should -BeOfType System.IO.FileInfo
        Test-Path -LiteralPath $path | Should -BeTrue

        $rtf = Get-OfficeRtf -Path $path
        $rtf.GetType().FullName | Should -Be 'OfficeIMO.Rtf.RtfReadResult'
        $rtf.Document.Paragraphs.Count | Should -Be 2
        $rtf.Document.Paragraphs[0].ToPlainText() | Should -Be 'Summary'
        $rtf.ToRtfLossless() | Should -Match 'Ready for review'
    }

    It 'updates RTF text losslessly and can append paragraphs and metadata' {
        $sourcePath = Join-Path $TestDrive 'Draft.rtf'
        $outputPath = Join-Path $TestDrive 'Final.rtf'
        New-OfficeRtf -Path $sourcePath -Text 'Draft summary' | Out-Null

        $file = Update-OfficeRtfText -Path $sourcePath -OutputPath $outputPath -OldText Draft -NewText Final -AppendParagraph 'Reviewed' -DocumentProperty @{ Title = 'Final report' } -UserProperty @{ Client = 'Contoso' } -DocumentVariable @{ Stage = 'Final' } -PassThru

        $file | Should -BeOfType System.IO.FileInfo
        $rtf = Get-OfficeRtf -Path $outputPath
        ($rtf.Document.Paragraphs | ForEach-Object ToPlainText) -join "`n" | Should -Match 'Final summary'
        ($rtf.Document.Paragraphs | ForEach-Object ToPlainText) -join "`n" | Should -Match 'Reviewed'
        $rtf.ToRtfLossless() | Should -Match 'Final report'
        $rtf.ToRtfLossless() | Should -Match 'Contoso'
        $rtf.ToRtfLossless() | Should -Match 'Stage'
    }
}
