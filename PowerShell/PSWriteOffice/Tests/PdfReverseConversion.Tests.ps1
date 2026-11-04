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
        New-OfficePdf -Path $pdfPath -Force {
            Add-OfficePdfHeading -Text 'Quarterly results' -Level 1
            Add-OfficePdfParagraph -Text 'Revenue improved in the current quarter.'
            Add-OfficePdfTable -InputObject @(
                [pscustomobject]@{ Region = 'North'; Revenue = 1250; Active = $true }
                [pscustomobject]@{ Region = 'South'; Revenue = 980; Active = $false }
                [pscustomobject]@{ Region = 'West'; Revenue = 1430; Active = $true }
            )
        } | Out-Null
    }

    It 'reconstructs Word content, Excel tables, and PowerPoint slides from a real PDF' {
        $wordPath = Join-Path $TestDrive 'source.docx'
        $excelPath = Join-Path $TestDrive 'source.xlsx'
        $powerPointPath = Join-Path $TestDrive 'source.pptx'

        $wordReport = Convert-OfficeDocument -Path $pdfPath -To Word -OutputPath $wordPath -Options (New-OfficePdfWordOptions -NoParagraphs:$false -NoHeadings:$false)
        $excelReport = Convert-OfficeDocument -Path $pdfPath -To Excel -OutputPath $excelPath -Options (New-OfficePdfExcelOptions -Culture en-US -NoNumericConversion:$false -ContinuationGeometryTolerance '4pt')
        $powerPointOptions = New-OfficePdfPowerPointOptions -Page 1 -Dpi 96.5
        $powerPointOptions.Dpi | Should -Be 96.5
        $powerPointReport = Convert-OfficeDocument -Path $pdfPath -To PowerPoint -OutputPath $powerPointPath -Options $powerPointOptions

        $wordReport.GetType().FullName | Should -Be 'OfficeIMO.OfficeConversionReport'
        $excelReport.GetType().FullName | Should -Be 'OfficeIMO.OfficeConversionReport'
        $powerPointReport.GetType().FullName | Should -Be 'OfficeIMO.OfficeConversionReport'
        Test-Path -LiteralPath $wordPath | Should -BeTrue
        Test-Path -LiteralPath $excelPath | Should -BeTrue
        Test-Path -LiteralPath $powerPointPath | Should -BeTrue

        $wordText = (Get-OfficeWordText -Path $wordPath | ForEach-Object Text) -join ' '
        $wordText | Should -Match 'Quarterly results'
        $wordText | Should -Match 'Revenue improved'
        $wordDocument = [OfficeIMO.Word.WordDocument]::Load($wordPath)
        try { $wordDocument.Tables[0].Rows[0].RepeatHeaderRowAtTheTopOfEachPage | Should -BeTrue }
        finally { $wordDocument.Dispose() }

        $excelReport.ItemsWritten | Should -BeGreaterThan 0
        $rows = @(Import-OfficeExcel -Path $excelPath -Sheet 1)
        $rows | Should -HaveCount 3
        $rows[0].Region | Should -Be 'North'
        $rows[0].Revenue | Should -Be 1250

        $presentation = Open-OfficePowerPoint -LiteralPath $powerPointPath
        try {
            $slides = @($presentation | Get-OfficePowerPointSlide -Summary)
            $slides | Should -HaveCount 1
            $slides[0].Title | Should -Be 'Quarterly results'
            $slides[0].TableCount | Should -BeGreaterThan 0
        } finally {
            $presentation | Close-OfficePowerPoint
        }
    }

    It 'preserves the tables-only profile unless an inverse switch explicitly restores a feature' {
        $tablesPath = Join-Path $TestDrive 'tables-only.docx'
        $headingsPath = Join-Path $TestDrive 'tables-and-headings.docx'
        Convert-OfficeDocument -Path $pdfPath -To Word -OutputPath $tablesPath -Options (New-OfficePdfWordOptions -TablesOnly) | Out-Null
        Convert-OfficeDocument -Path $pdfPath -To Word -OutputPath $headingsPath -Options (New-OfficePdfWordOptions -TablesOnly -NoHeadings:$false) | Out-Null
        $tablesText = (Get-OfficeWordText -Path $tablesPath | ForEach-Object Text) -join ' '
        $tablesDocument = [OfficeIMO.Word.WordDocument]::Load($tablesPath)
        try {
            $tablesDocument.Tables.Count | Should -BeGreaterThan 0
            ($tablesDocument.Tables[0].Rows[1].Cells[0].Paragraphs.Text -join ' ') | Should -Be 'North'
        } finally { $tablesDocument.Dispose() }
        $tablesText | Should -Not -Match 'Quarterly results'
        $tablesText | Should -Not -Match 'Revenue improved'
        $headingsText = (Get-OfficeWordText -Path $headingsPath | ForEach-Object Text) -join ' '
        $headingsText | Should -Match 'Quarterly results'
        $headingsText | Should -Not -Match 'Revenue improved'
    }

    It 'honors WhatIf, extension validation, and explicit overwrite' {
        $wordPath = Join-Path $TestDrive 'what-if.docx'
        Convert-OfficeDocument -To Word -Path $pdfPath -OutputPath $wordPath -WhatIf | Out-Null
        Test-Path -LiteralPath $wordPath | Should -BeFalse

        $wrongPath = Join-Path $TestDrive 'wrong.pdf'
        try {
            Convert-OfficeDocument -To Word -Path $pdfPath -OutputPath $wrongPath -ErrorAction Stop
            throw 'Expected invalid-extension failure.'
        } catch {
            $_.Exception.Message | Should -Match 'To and the OutputPath extension must name the same target format'
            $_.CategoryInfo.Category | Should -Be 'InvalidArgument'
        }

        Convert-OfficeDocument -To Word -Path $pdfPath -OutputPath $wordPath | Out-Null
        $before = [Convert]::ToBase64String([IO.File]::ReadAllBytes($wordPath))
        $beforeTimestamp = (Get-Item -LiteralPath $wordPath).LastWriteTimeUtc
        { Convert-OfficeDocument -To Word -Path $pdfPath -OutputPath $wordPath -NoClobber -ErrorAction Stop } |
            Should -Throw
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($wordPath)) | Should -Be $before
        (Get-Item -LiteralPath $wordPath).LastWriteTimeUtc | Should -Be $beforeTimestamp
        Convert-OfficeDocument -To Word -Path $pdfPath -OutputPath $wordPath | Out-Null
        Test-Path -LiteralPath $wordPath | Should -BeTrue

        $blockedDirectory = Join-Path $TestDrive 'not-a-directory'
        Set-Content -LiteralPath $blockedDirectory -Value 'file'
        $unwritableOutput = Join-Path $blockedDirectory 'result.docx'
        try {
            Convert-OfficeDocument -To Word -Path $pdfPath -OutputPath $unwritableOutput -ErrorAction Stop
            throw 'Expected output-write failure.'
        } catch {
            $_.Exception | Should -Not -BeNullOrEmpty
            $_.Exception.Message | Should -Not -Be 'Expected output-write failure.'
            $_.CategoryInfo.Category | Should -Be 'WriteError'
            $_.TargetObject | Should -Be $unwritableOutput
            Test-Path -LiteralPath $unwritableOutput | Should -BeFalse
        }
    }
}

Describe 'Protected-content capability discovery' {
    It 'returns filterable typed rows and deterministic catalog output' {
        $all = @(Get-OfficeDocumentCapability -Protection)
        $all.Count | Should -BeGreaterThan 0
        $all[0].GetType().FullName | Should -Be 'OfficeIMO.Security.OfficeProtectionCapability'

        $pdf = @(Get-OfficeDocumentCapability -Protection -Format PDF)
        $pdf | Should -HaveCount 1
        $pdf[0].Id | Should -Be 'pdf-password'

        $incomplete = @(Get-OfficeDocumentCapability -Protection -IncompleteOnly)
        $incomplete.Count | Should -BeGreaterThan 0
        $incomplete.Count | Should -BeLessThan $all.Count

        $catalog = Get-OfficeDocumentCapability -Protection -AsJson | ConvertFrom-Json
        $catalog.id | Should -Be 'OfficeIMO.ProtectedContent'
        $catalog.schemaVersion | Should -Be 2
        @($catalog.capabilities).Count | Should -Be $all.Count
        $catalog.capabilities[0].inspect | Should -BeIn @('Supported', 'Unsupported', 'NotApplicable', 'Partial')
        $catalog.capabilities[0].PSObject.Properties.Name | Should -Contain 'unsupportedOperations'
    }

    It 'rejects ambiguous text output requests' {
        { Get-OfficeDocumentCapability -Protection -AsJson -AsMarkdown -ErrorAction Stop } |
            Should -Throw '*Specify only one of -AsJson or -AsMarkdown*'
        { Get-OfficeDocumentCapability -Protection -AsJson -Format PDF -ErrorAction Stop } |
            Should -Throw '*cannot be combined with row filters*'
    }
}
