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

        $docFile = ConvertTo-OfficeWordDocument -Path $docxPath -OutputPath $docPath -PassThru
        $docFile | Should -BeOfType System.IO.FileInfo
        Test-Path -LiteralPath $docPath | Should -BeTrue
        [BitConverter]::ToString([IO.File]::ReadAllBytes($docPath)[0..3]) | Should -Be 'D0-CF-11-E0'

        ConvertTo-OfficeWordDocument -Path $docPath -OutputPath $roundTripPath
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

        $xlsFile = ConvertTo-OfficeExcelWorkbook -Path $xlsxPath -OutputPath $xlsPath -PassThru
        $xlsFile | Should -BeOfType System.IO.FileInfo
        Test-Path -LiteralPath $xlsPath | Should -BeTrue
        [BitConverter]::ToString([IO.File]::ReadAllBytes($xlsPath)[0..3]) | Should -Be 'D0-CF-11-E0'

        ConvertTo-OfficeExcelWorkbook -Path $xlsPath -OutputPath $roundTripPath
        $rows = @(Import-OfficeExcel -Path $roundTripPath -WorksheetName 'Data')
        $rows | Should -HaveCount 1
        $rows[0].Name | Should -Be 'Ada'
        [int] $rows[0].Value | Should -Be 42
    }

    It 'does not overwrite conversion outputs unless Force is used' {
        $sourcePath = Join-Path $TestDrive 'source.docx'
        $targetPath = Join-Path $TestDrive 'target.docx'

        New-OfficeWord -Path $sourcePath {
            Add-OfficeWordParagraph -Text 'Overwrite source'
        } | Out-Null
        New-OfficeWord -Path $targetPath {
            Add-OfficeWordParagraph -Text 'Keep me'
        } | Out-Null

        { ConvertTo-OfficeWordDocument -Path $sourcePath -OutputPath $targetPath -ErrorAction Stop } |
            Should -Throw '*already exists*'

        ConvertTo-OfficeWordDocument -Path $sourcePath -OutputPath $targetPath -Force
        $paragraphs = @(Get-OfficeWordParagraph -Path $targetPath)
        $paragraphs.Text | Should -Contain 'Overwrite source'
    }
}
