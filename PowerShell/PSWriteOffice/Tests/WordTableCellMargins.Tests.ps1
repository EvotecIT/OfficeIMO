BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $manifest -Global -ErrorAction Stop
}

Describe 'Word table cell padding' {
    It 'rejects an invalid margin before changing the cell' -ForEach @(
        @{ Value = '-0.1cm' }, @{ Value = '58cm' }, @{ Value = [double]::NaN }
    ) {
        New-OfficeWord -Path (Join-Path $TestDrive ("Invalid-" + [guid]::NewGuid().ToString("N") + ".docx")) {
            $table = WordTable -InputObject @([pscustomobject]@{ Site = 'Warsaw' }) -PassThru
            $cell = $table | Get-OfficeWordTableCell -Row 2 -Column 1
            { $cell | Set-OfficeWordTableCell -Text 'Changed' -MarginTop $Value } | Should -Throw
            $cell.Paragraphs[0].Text | Should -Be 'Warsaw'
        }
    }

    It 'persists independent margins through the cell pipeline without changing text' {
        $path = Join-Path $TestDrive 'Padded.docx'
        New-OfficeWord -Path $path {
            $table = WordTable -InputObject @([pscustomobject]@{ Site = 'Warsaw'; Status = 'Ready' }) -PassThru
            $table | Get-OfficeWordTableCell | Set-OfficeWordTableCell -MarginTop 0.12cm -MarginBottom 0.14cm -MarginLeft 0.16cm -MarginRight 0.18cm
        }
        $doc = Open-OfficeWord -Path $path
        try {
            $cells = @($doc.Tables[0] | Get-OfficeWordTableCell)
            $cells.Count | Should -Be 4
            foreach ($cell in $cells) {
                [math]::Abs($cell.MarginTop.ToCentimeters() - 0.12) | Should -BeLessThan 0.002
                [math]::Abs($cell.MarginBottom.ToCentimeters() - 0.14) | Should -BeLessThan 0.002
                [math]::Abs($cell.MarginLeft.ToCentimeters() - 0.16) | Should -BeLessThan 0.002
                [math]::Abs($cell.MarginRight.ToCentimeters() - 0.18) | Should -BeLessThan 0.002
            }
            $cells[2].Paragraphs[0].Text | Should -Be 'Warsaw'
            @($doc.ValidateDocument()).Count | Should -Be 0
        } finally { $doc | Close-OfficeWord }
    }

    It 'preserves unbound margins and removes only an explicitly cleared override' {
        $path = Join-Path $TestDrive 'Cleared.docx'
        New-OfficeWord -Path $path {
            $table = WordTable -InputObject @([pscustomobject]@{ Site = 'Warsaw' }) -PassThru
            $cell = $table | Get-OfficeWordTableCell -Row 2 -Column 1
            $cell | Set-OfficeWordTableCell -MarginTop 0.12cm -MarginLeft 0.16cm
            $cell | Set-OfficeWordTableCell -Text 'Dublin' -MarginTop $null
        }
        $doc = Open-OfficeWord -Path $path
        try {
            $cell = $doc.Tables[0].Rows[1].Cells[0]
            $cell.MarginTop | Should -BeNullOrEmpty
            [math]::Abs($cell.MarginLeft.ToCentimeters() - 0.16) | Should -BeLessThan 0.002
            $cell.Paragraphs[0].Text | Should -Be 'Dublin'
        } finally { $doc | Close-OfficeWord }
    }
}
