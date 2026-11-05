BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'PDF forms and stamps' {
    It 'builds, reads, fills, and flattens PDF form fields' {
        $formPath = Join-Path $TestDrive 'form.pdf'
        New-OfficePdf -Path $formPath {
            PdfHeading 'Customer intake'
            PdfParagraph 'Customer name'
            PdfFormField -Name 'CustomerName' -Type Text
            PdfParagraph 'Plan'
            PdfFormField -Name 'Plan' -Type Choice -Option 'Standard', 'Premium' -Value 'Standard'
            PdfFormField -Name 'Regions' -Type MultiSelectChoice -Option 'EU', 'US', 'APAC' -Value 'EU', 'US'
        } | Out-Null

        $fields = Get-OfficePdfFormField -Path $formPath
        $fields.Name | Should -Contain 'CustomerName'
        $fields.Name | Should -Contain 'Plan'
        $fields.Name | Should -Contain 'Regions'

        $live = Open-OfficePdf -LiteralPath $formPath
        try {
            $preview = @($live | Set-OfficePdfFormField -Value @{ CustomerName = 'Preview' } -PassThru -WhatIf)
            try { $preview.Count | Should -Be 0 }
            finally { foreach ($result in $preview) { if (-not [object]::ReferenceEquals($result, $live)) { $result.Dispose() } } }
        } finally { $live.Dispose() }

        $contextPath = Join-Path $TestDrive 'context-filled.pdf'
        Open-OfficePdf -LiteralPath $formPath -OutputPath $contextPath {
            Set-OfficePdfFormField -Value @{ CustomerName = 'After' }
        } | Out-Null
        (Get-OfficePdfFormField -LiteralPath $contextPath -Name CustomerName).Value | Should -Be 'After'

        $filledPath = Join-Path $TestDrive 'filled.pdf'
        Set-OfficePdfFormField -Path $formPath -OutputPath $filledPath -Value @{
            CustomerName = 'Alice Example'
            Plan = 'Premium'
            Regions = 'EU'
        } -Flatten -PassThru | Should -BeOfType System.IO.FileInfo

        $preflight = Get-OfficePdfDiagnostic -Type Preflight -Path $filledPath
        $preflight.CanRead | Should -BeTrue
        ((Get-OfficeDocumentInfo -Detailed -Path $filledPath).Details).FormFieldCount | Should -Be 0
    }

    It 'creates the output directory when filling PDF forms' {
        $formPath = Join-Path $TestDrive 'nested-form-source.pdf'
        New-OfficePdf -Path $formPath {
            PdfHeading 'Customer intake'
            PdfFormField -Name 'CustomerName' -Type Text
        } | Out-Null

        $filledPath = Join-Path $TestDrive 'nested\filled.pdf'
        Set-OfficePdfFormField -Path $formPath -OutputPath $filledPath -Value @{
            CustomerName = 'Alice Example'
        } -PassThru | Should -BeOfType System.IO.FileInfo

        Test-Path $filledPath | Should -BeTrue
        (Get-OfficePdfDiagnostic -Type Preflight -Path $filledPath).CanRead | Should -BeTrue
    }

    It 'writes deterministic form appearances without requiring viewer regeneration' {
        $formPath = Join-Path $TestDrive 'appearance-source.pdf'
        New-OfficePdf -Path $formPath {
            PdfHeading 'Customer intake'
            PdfFormField -Name 'CustomerName' -Type Text
        } | Out-Null

        $filledPath = Join-Path $TestDrive 'appearance-filled.pdf'
        Set-OfficePdfFormField -Path $formPath -OutputPath $filledPath -Value @{
            CustomerName = 'Alice Example'
        } -PassThru | Should -BeOfType System.IO.FileInfo

        $info = (Get-OfficeDocumentInfo -Detailed -Path $filledPath).Details
        $raw = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($filledPath))

        $info.AcroFormNeedAppearances | Should -BeFalse
        $raw | Should -Match '/AP << /N'
    }

    It 'can keep NeedAppearances for legacy PDF viewers' {
        $formPath = Join-Path $TestDrive 'legacy-appearance-source.pdf'
        New-OfficePdf -Path $formPath {
            PdfHeading 'Customer intake'
            PdfFormField -Name 'CustomerName' -Type Text
        } | Out-Null

        $filledPath = Join-Path $TestDrive 'legacy-appearance-filled.pdf'
        Set-OfficePdfFormField -Path $formPath -OutputPath $filledPath -Value @{
            CustomerName = 'Alice Example'
        } -KeepNeedAppearances -PassThru | Should -BeOfType System.IO.FileInfo

        ((Get-OfficeDocumentInfo -Detailed -Path $filledPath).Details).AcroFormNeedAppearances | Should -BeTrue
    }

    It 'updates existing PDF metadata and adds extractable text stamps' {
        $sourcePath = Join-Path $TestDrive 'source.pdf'
        New-OfficePdf -Path $sourcePath {
            PdfHeading 'Invoice'
            PdfParagraph 'Original body'
        } | Out-Null

        $metadataPath = Join-Path $TestDrive 'metadata.pdf'
        Set-OfficePdfMetadata -Path $sourcePath -OutputPath $metadataPath -Title 'Stamped Invoice' -Author 'PSWriteOffice' -PassThru |
            Should -BeOfType System.IO.FileInfo

        $metadata = ((Get-OfficeDocumentInfo -Detailed -Path $metadataPath).Details).Metadata
        $metadata.Title | Should -Be 'Stamped Invoice'
        $metadata.Author | Should -Be 'PSWriteOffice'

        $stampedPath = Join-Path $TestDrive 'stamped.pdf'
        Add-OfficePdfStamp -Path $metadataPath -OutputPath $stampedPath -Text 'APPROVED' -X 72 -Y 72 -Size 18 -Color '#008000' -PassThru |
            Should -BeOfType System.IO.FileInfo

        $text = Read-OfficeDocument -As Text -Path $stampedPath
        $text | Should -Match 'Invoice'
        $text | Should -Match 'APPROVED'
    }

    It 'creates output directories for PDF metadata and stamps' {
        $sourcePath = Join-Path $TestDrive 'source-nested.pdf'
        New-OfficePdf -Path $sourcePath {
            PdfHeading 'Invoice'
            PdfParagraph 'Original body'
        } | Out-Null

        $metadataPath = Join-Path $TestDrive 'metadata\out.pdf'
        Set-OfficePdfMetadata -Path $sourcePath -OutputPath $metadataPath -Title 'Nested Metadata' -PassThru |
            Should -BeOfType System.IO.FileInfo

        $stampedPath = Join-Path $TestDrive 'stamps\approved.pdf'
        Add-OfficePdfStamp -Path $metadataPath -OutputPath $stampedPath -Text 'APPROVED' -X 72 -Y 72 -PassThru |
            Should -BeOfType System.IO.FileInfo

        Test-Path $metadataPath | Should -BeTrue
        Test-Path $stampedPath | Should -BeTrue
        ((Get-OfficeDocumentInfo -Detailed -Path $metadataPath).Details).Metadata.Title | Should -Be 'Nested Metadata'
        Read-OfficeDocument -As Text -Path $stampedPath | Should -Match 'APPROVED'
    }

    It 'converts PDF forms to flat PDFs with an approved verb' {
        $formPath = Join-Path $TestDrive 'flat-source.pdf'
        New-OfficePdf -Path $formPath {
            PdfHeading 'Approval'
            PdfFormField -Name 'ApprovedBy' -Type Text -Value 'Reviewer'
        } | Out-Null

        $flatPath = Join-Path $TestDrive 'flat.pdf'
        Update-OfficePdf -FlattenForm -Path $formPath -OutputPath $flatPath | Should -BeOfType System.IO.FileInfo

        ((Get-OfficeDocumentInfo -Detailed -Path $flatPath).Details).FormFieldCount | Should -Be 0
        (Get-OfficePdfDiagnostic -Type Preflight -Path $flatPath).CanRead | Should -BeTrue
    }

    It 'creates the output directory when converting forms to flat PDFs' {
        $formPath = Join-Path $TestDrive 'flat-nested-source.pdf'
        New-OfficePdf -Path $formPath {
            PdfHeading 'Approval'
            PdfFormField -Name 'ApprovedBy' -Type Text -Value 'Reviewer'
        } | Out-Null

        $flatPath = Join-Path $TestDrive 'flat\nested.pdf'
        Update-OfficePdf -FlattenForm -Path $formPath -OutputPath $flatPath | Should -BeOfType System.IO.FileInfo

        Test-Path $flatPath | Should -BeTrue
        ((Get-OfficeDocumentInfo -Detailed -Path $flatPath).Details).FormFieldCount | Should -Be 0
    }
}
