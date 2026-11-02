BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $moduleManifest -Global -ErrorAction Stop
}

Describe 'Conversion capability discovery' {
    It 'returns native fidelity evidence and intersects exact extensions and IDs' {
        $route = @(Get-OfficeDocumentCapability -Conversion -Id DOCX-PDF -From 'DOCX' -To '.PDF')
        $route | Should -HaveCount 1
        $route[0].Id | Should -Be 'docx-pdf'
        $route[0].PackageId | Should -Be 'OfficeIMO.Word.Pdf'
        $route[0].KnownLimitations | Should -Not -BeNullOrEmpty
        $route[0].SupportEvidence | Should -Not -BeNullOrEmpty
        $route[0].TextFormattingContract | Should -Not -BeNullOrEmpty
        @(Get-OfficeDocumentCapability -Conversion -Id docx-pdf -To docx) | Should -HaveCount 0
        @(Get-OfficeDocumentCapability -Conversion -Id no-such-route) | Should -HaveCount 0
        $families = @(Get-OfficeDocumentCapability -Conversion -From Word -To Pdf)
        $families.Id | Should -Contain 'docx-pdf'
        { Get-OfficeDocumentCapability -Conversion -From no-such-format -ErrorAction Stop } | Should -Throw '*supported format*'
    }

    It 'reports installed executable routes without describing unavailable converters as executable' {
        $executable = @(Get-OfficeDocumentCapability -Conversion -Executable)
        $executable.Id | Should -Contain 'docx-pdf'
        $executable.Id | Should -Contain 'pdf-docx'
        $executable.ConversionSupported | Should -Not -Contain $false
        $executable.ConversionCommand | Should -Not -Contain $null
        $executable.ConversionCommand | Select-Object -Unique | Should -Be 'Convert-OfficeDocument'
        $pdf = @(Get-OfficeDocumentCapability -Conversion -Executable -To Pdf)
        $pdf.Id | Should -Contain 'docx-pdf'
        $pdf.Id | Should -Not -Contain 'pdf-docx'
        @($pdf | Where-Object TargetExtension -NE '.pdf') | Should -HaveCount 0
    }

    It 'accepts pipeline route IDs without mutating canonical rows on repeated queries' {
        $rows = @([pscustomobject]@{ Id = 'docx-pdf' }, [pscustomobject]@{ Id = 'pdf-docx' }) |
            Get-OfficeDocumentCapability -Conversion
        @($rows) | Should -HaveCount 2
        @(Get-OfficeDocumentCapability -Conversion -Id docx-pdf) | Should -HaveCount 1
        @(Get-OfficeDocumentCapability -Conversion -Id docx-pdf) | Should -HaveCount 1
    }
}
