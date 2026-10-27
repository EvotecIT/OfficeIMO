BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $moduleManifest -Global -ErrorAction Stop
}
Describe 'Foundation output views' {
    It 'registers compact report, diagnostic, length, theme and image views' {
        foreach ($name in 'OfficeIMO.OfficeConversionReport','OfficeIMO.OfficeDiagnostic','OfficeIMO.OfficeLength','OfficeIMO.OfficeTheme','OfficeIMO.Drawing.OfficeImageExportResult') {
            @(Get-FormatData -TypeName $name).Count | Should -BeGreaterThan 0
        }
        [OfficeIMO.OfficeLength]::Parse('5cm') | Out-String | Should -Match '141.7323pt'
        [OfficeIMO.OfficeTheme]::Modern | Out-String | Should -Match 'Modern'
        $report = [OfficeIMO.OfficeConversionReport]::new('in.docx','Pdf','out.pdf',2)
        $report | Out-String | Should -Match 'out.pdf'
        $diagnostic = [OfficeIMO.OfficeDiagnostic]::new('chart-flattened','Warning','Chart flattened')
        $diagnostic | Out-String | Should -Match 'Chart flattened'
    }
    It 'shows document, paragraph, sheet, slide, table and image summaries without binary content' {
        $word = [OfficeIMO.Word.WordDocument]::Create()
        $excel = [OfficeIMO.Excel.ExcelDocument]::Create()
        $deck = [OfficeIMO.PowerPoint.PowerPointDocument]::Create()
        try {
            $paragraph = $word.AddParagraph('Visible paragraph')
            $sheet = $excel.AddWorksheet('Visible sheet')
            $slide = $deck.AddSlide()
            $table = $sheet.Table('A1:B2')
            $png = [OfficeIMO.Drawing.OfficePngWriter]::EncodeRgba(1,1,[byte[]]@(255,0,0,255))
            $image = [OfficeIMO.Drawing.OfficeImageExportResult]::new('Png',1,1,$png,'Small image')
            foreach ($value in @($word,$excel,$deck,$paragraph,$sheet,$slide,$table,$image)) {
                $output = $value | Out-String
                $output | Should -Not -Match '\bBytes\b|System.Byte\[\]'
                $output.Trim().Length | Should -BeGreaterThan 0
            }
            $paragraph | Out-String | Should -Match 'Visible paragraph'
            $sheet | Out-String | Should -Match 'Visible sheet'
            $table | Out-String | Should -Match 'A1:B2'
            $image | Out-String | Should -Match 'Small image'
        } finally { $word.Dispose(); $excel.Dispose(); $deck.Dispose() }
    }
}
