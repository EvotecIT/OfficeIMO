BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'PDF reverse conversion workflows' {
    BeforeEach {
        $pdfPath = Join-Path $TestDrive 'source.pdf'
        New-OfficePdf -Path $pdfPath {
            Add-OfficePdfHeading -Text 'Quarterly results' -Level 1
            Add-OfficePdfParagraph -Text 'Revenue improved in the current quarter.'
            Add-OfficePdfTable -InputObject @(
                [pscustomobject]@{ Region = 'North'; Revenue = 1250; Active = $true }
                [pscustomobject]@{ Region = 'South'; Revenue = 980; Active = $false }
                [pscustomobject]@{ Region = 'West'; Revenue = 1430; Active = $true }
            )
        } | Out-Null
    }

    It 'exports the three editable PDF conversion commands' {
        foreach ($command in 'ConvertTo-OfficePdfWord', 'ConvertTo-OfficePdfExcel', 'ConvertTo-OfficePdfPowerPoint') {
            Get-Command $command -ErrorAction Stop | Should -Not -BeNullOrEmpty
        }
    }

    It 'reconstructs Word content, Excel tables, and PowerPoint slides from a real PDF' {
        $wordPath = Join-Path $TestDrive 'source.docx'
        $excelPath = Join-Path $TestDrive 'source.xlsx'
        $powerPointPath = Join-Path $TestDrive 'source.pptx'

        $wordReport = ConvertTo-OfficePdfWord -Path $pdfPath -OutputPath $wordPath -PassThruReport
        $excelReport = ConvertTo-OfficePdfExcel -Path $pdfPath -OutputPath $excelPath -PassThruReport
        $powerPointReport = ConvertTo-OfficePdfPowerPoint -Path $pdfPath -OutputPath $powerPointPath -PassThruReport

        $wordReport.GetType().FullName | Should -Be 'OfficeIMO.Word.Pdf.PdfWordConversionReport'
        $excelReport.GetType().FullName | Should -Be 'OfficeIMO.Excel.Pdf.PdfExcelTableImportReport'
        $powerPointReport.GetType().FullName | Should -Be 'OfficeIMO.PowerPoint.Pdf.PdfPowerPointConversionReport'
        Test-Path -LiteralPath $wordPath | Should -BeTrue
        Test-Path -LiteralPath $excelPath | Should -BeTrue
        Test-Path -LiteralPath $powerPointPath | Should -BeTrue

        $wordText = (Get-OfficeWordText -Path $wordPath | ForEach-Object Text) -join ' '
        $wordText | Should -Match 'Quarterly results'
        $wordText | Should -Match 'Revenue improved'

        $excelReport.Entries.Count | Should -BeGreaterThan 0
        $excelSummary = Get-OfficeExcelSummary -Path $excelPath -IncludeSheets
        $rows = @(Import-OfficeExcel -Path $excelPath -WorksheetName $excelSummary.Sheets[0].Name)
        $rows | Should -HaveCount 3
        $rows[0].Region | Should -Be 'North'
        $rows[0].Revenue | Should -Be 1250

        $presentation = Get-OfficePowerPoint -Path $powerPointPath
        try {
            $slides = @($presentation | Get-OfficePowerPointSlideSummary)
            $slides | Should -HaveCount 1
            $slides[0].Title | Should -Be 'Quarterly results'
            $slides[0].TableCount | Should -BeGreaterThan 0
        } finally {
            $presentation | Close-OfficePowerPoint
        }
    }

    It 'honors WhatIf, extension validation, and explicit overwrite' {
        $wordPath = Join-Path $TestDrive 'what-if.docx'
        ConvertTo-OfficePdfWord -Path $pdfPath -OutputPath $wordPath -WhatIf | Out-Null
        Test-Path -LiteralPath $wordPath | Should -BeFalse

        $wrongPath = Join-Path $TestDrive 'wrong.pdf'
        try {
            ConvertTo-OfficePdfWord -Path $pdfPath -OutputPath $wrongPath -ErrorAction Stop
            throw 'Expected invalid-extension failure.'
        } catch {
            $_.Exception.Message | Should -Match 'must use the .docx extension'
            $_.CategoryInfo.Category | Should -Be 'InvalidArgument'
        }

        ConvertTo-OfficePdfWord -Path $pdfPath -OutputPath $wordPath | Out-Null
        try {
            ConvertTo-OfficePdfWord -Path $pdfPath -OutputPath $wordPath -ErrorAction Stop
            throw 'Expected existing-output failure.'
        } catch {
            $_.Exception.Message | Should -Match 'Use -Force to overwrite it'
            $_.CategoryInfo.Category | Should -Be 'ResourceExists'
            $_.TargetObject | Should -Be $wordPath
        }
        ConvertTo-OfficePdfWord -Path $pdfPath -OutputPath $wordPath -Force | Out-Null
        Test-Path -LiteralPath $wordPath | Should -BeTrue

        $blockedDirectory = Join-Path $TestDrive 'not-a-directory'
        Set-Content -LiteralPath $blockedDirectory -Value 'file'
        $unwritableOutput = Join-Path $blockedDirectory 'result.docx'
        try {
            ConvertTo-OfficePdfWord -Path $pdfPath -OutputPath $unwritableOutput -ErrorAction Stop
            throw 'Expected output-write failure.'
        } catch {
            $_.CategoryInfo.Category | Should -Be 'WriteError'
            $_.TargetObject | Should -Be $unwritableOutput
        }
    }
}

Describe 'Protected-content capability discovery' {
    It 'returns filterable typed rows and deterministic catalog output' {
        $all = @(Get-OfficeProtectionCapability)
        $all.Count | Should -BeGreaterThan 0
        $all[0].GetType().FullName | Should -Be 'OfficeIMO.Security.OfficeProtectionCapability'

        $pdf = @(Get-OfficeProtectionCapability -Format PDF)
        $pdf | Should -HaveCount 1
        $pdf[0].Id | Should -Be 'pdf-password'

        $incomplete = @(Get-OfficeProtectionCapability -IncompleteOnly)
        $incomplete.Count | Should -BeGreaterThan 0
        $incomplete.Count | Should -BeLessThan $all.Count

        $catalog = Get-OfficeProtectionCapability -AsJson | ConvertFrom-Json
        $catalog.id | Should -Be 'OfficeIMO.ProtectedContent'
        $catalog.schemaVersion | Should -Be 2
        @($catalog.capabilities).Count | Should -Be $all.Count
        $catalog.capabilities[0].inspect | Should -BeIn @('Supported', 'Unsupported', 'NotApplicable', 'Partial')
        $catalog.capabilities[0].PSObject.Properties.Name | Should -Contain 'unsupportedOperations'
    }

    It 'rejects ambiguous text output requests' {
        { Get-OfficeProtectionCapability -AsJson -AsMarkdown -ErrorAction Stop } |
            Should -Throw '*Specify only one of -AsJson or -AsMarkdown*'
        { Get-OfficeProtectionCapability -AsJson -Format PDF -ErrorAction Stop } |
            Should -Throw '*cannot be combined with row filters*'
    }
}
