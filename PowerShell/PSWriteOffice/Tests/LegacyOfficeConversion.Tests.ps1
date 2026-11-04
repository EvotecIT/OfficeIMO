BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Legacy Office conversion cmdlets' {
    It 'converts Word documents between docx and doc' {
        $docxPath = Join-Path $TestDrive 'source.docx'
        $docPath = Join-Path $TestDrive 'converted.doc'
        $roundTripPath = Join-Path $TestDrive 'roundtrip.docx'

        New-OfficeWord -Path $docxPath {
            Add-OfficeWordParagraph -Text 'Legacy Word conversion'
        } | Out-Null

        $docFile = Convert-OfficeDocument -Path $docxPath -OutputPath $docPath
        $docFile | Should -BeOfType OfficeIMO.OfficeConversionReport
        Test-Path -LiteralPath $docPath | Should -BeTrue
        [BitConverter]::ToString([IO.File]::ReadAllBytes($docPath)[0..3]) | Should -Be 'D0-CF-11-E0'

        Convert-OfficeDocument -Path $docPath -OutputPath $roundTripPath
        $paragraphs = @(Get-OfficeWordParagraph -Path $roundTripPath)
        $paragraphs.Text | Should -Contain 'Legacy Word conversion'
    }

    It 'converts Excel workbooks between xlsx and xls' {
        $xlsxPath = Join-Path $TestDrive 'source.xlsx'
        $xlsPath = Join-Path $TestDrive 'converted.xls'
        $roundTripPath = Join-Path $TestDrive 'roundtrip.xlsx'

        New-OfficeExcel -Path $xlsxPath {
            ExcelSheet 'Data' {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelRow -Row 2 -Values 'Ada', 42
            }
        } | Out-Null

        $xlsFile = Convert-OfficeDocument -Path $xlsxPath -OutputPath $xlsPath
        $xlsFile | Should -BeOfType OfficeIMO.OfficeConversionReport
        Test-Path -LiteralPath $xlsPath | Should -BeTrue
        [BitConverter]::ToString([IO.File]::ReadAllBytes($xlsPath)[0..3]) | Should -Be 'D0-CF-11-E0'

        Convert-OfficeDocument -Path $xlsPath -OutputPath $roundTripPath
        $rows = @(Import-OfficeExcel -Path $roundTripPath -Sheet 'Data')
        $rows | Should -HaveCount 1
        $rows[0].Name | Should -Be 'Ada'
        [int] $rows[0].Value | Should -Be 42
    }

    It 'honors NoClobber before explicitly replacing an existing conversion output' {
        $sourcePath = Join-Path $TestDrive 'overwrite-source.docx'
        $existingPath = Join-Path $TestDrive 'existing-target.docx'
        $targetPath = Join-Path $TestDrive 'overwrite-target.doc'

        New-OfficeWord -Path $sourcePath {
            Add-OfficeWordParagraph -Text 'Overwrite source'
        } | Out-Null
        New-OfficeWord -Path $existingPath {
            Add-OfficeWordParagraph -Text 'Keep me'
        } | Out-Null

        Convert-OfficeDocument -Path $existingPath -OutputPath $targetPath | Out-Null
        $originalBytes = [IO.File]::ReadAllBytes($targetPath)

        $skipped = Convert-OfficeDocument -Path $sourcePath -OutputPath $targetPath -NoClobber -ErrorAction Stop
        $skipped.Skipped | Should -BeTrue
        $skipped.SkippedReason | Should -Be 'OutputExists'
        [IO.File]::ReadAllBytes($targetPath) | Should -Be $originalBytes

        Convert-OfficeDocument -Path $sourcePath -OutputPath $targetPath
        $paragraphs = @(Get-OfficeWordParagraph -Path $targetPath)
        $paragraphs.Text | Should -Contain 'Overwrite source'
    }
}
