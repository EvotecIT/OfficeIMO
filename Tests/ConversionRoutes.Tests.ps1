BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '../PSWriteOffice.psd1' }
    Import-Module $moduleManifest -Global -ErrorAction Stop
}

Describe 'Conversion route discovery' {
    It 'returns the owner contract and intersects extension and ID filters' {
        $route = @(Get-OfficeConversionRoute -Id DOCX-PDF -SourceExtension 'DOCX' -TargetExtension '.PDF')
        $route.Count | Should -Be 1
        $route[0].Id | Should -Be 'docx-pdf'
        $route[0].PackageId | Should -Be 'OfficeIMO.Word.Pdf'
        $route[0].KnownLimitations | Should -Not -BeNullOrEmpty
        $route[0].SupportEvidence | Should -Not -BeNullOrEmpty
        $route[0].TextFormattingContract | Should -Not -BeNullOrEmpty
        @(Get-OfficeConversionRoute -Id docx-pdf -TargetExtension docx).Count | Should -Be 0
        @(Get-OfficeConversionRoute -Id no-such-route).Count | Should -Be 0
    }

    It 'reports batch PDF availability for the current host without treating reverse routes as PDF export' {
        $word = Get-OfficeConversionRoute -Id docx-pdf
        $reverse = Get-OfficeConversionRoute -Id pdf-docx
        $modern = $PSVersionTable.PSEdition -eq 'Core'
        $word.PdfBatchSupported | Should -Be $modern
        $word.PdfBatchMinimumPowerShellVersion | Should -Be ([version]'7.4')
        $reverse.PdfBatchSupported | Should -BeFalse
        $reverse.PdfBatchCommand | Should -BeNullOrEmpty
        $batch = @(Get-OfficeConversionRoute -PdfBatch)
        if ($modern) {
            $batch.Id | Should -Contain 'docx-pdf'
            $batch.Id | Should -Not -Contain 'pdf-docx'
            @($batch | Where-Object TargetExtension -ne '.pdf').Count | Should -Be 0
            @($batch | Where-Object PdfBatchCommand -ne 'Export-OfficeDocumentPdf').Count | Should -Be 0
        } else { $batch.Count | Should -Be 0 }
    }

    It 'can be queried repeatedly and from pipeline IDs without modifying catalog rows' {
        $rows = @([pscustomobject]@{ Id = 'docx-pdf' }, [pscustomobject]@{ Id = 'pdf-docx' }) | Get-OfficeConversionRoute
        @($rows).Count | Should -Be 2
        @(Get-OfficeConversionRoute -Id docx-pdf).Count | Should -Be 1
        @(Get-OfficeConversionRoute -Id docx-pdf).Count | Should -Be 1
    }
}
