BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '..\PSWriteOffice.psd1' }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Excel encrypted lifecycle and conversion' {
    It 'round-trips encrypted workbooks through lifecycle cmdlets' {
        $path = Join-Path $TestDrive 'EncryptedExcel.xlsx'

        New-OfficeExcel -Path $path -Password 'secret' -SaveOptions @{ Preflight = $true } {
            Set-OfficeExcelExecutionPolicy -Mode Sequential -ParallelThreshold 5 -WorksheetValidation Always -Diagnostics
            Add-OfficeExcelSheet -Name 'Secure' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Encrypted value'
            }
        }

        { Get-ZipEntriesLocal -Path $path } | Should -Throw
        $doc = Get-OfficeExcel -Path $path -Password 'secret' -ReadOnly
        try {
            $doc.Sheets[0].Name | Should -Be 'Secure'
            $value = $null
            $doc.Sheets[0].TryGetCellText(1, 1, [ref] $value) | Should -BeTrue
            $value | Should -Be 'Encrypted value'
            $summary = Get-OfficeExcelSummary -Document $doc -IncludeSheets -IncludeSchema
            $summary.SheetCount | Should -Be 1
            $summary.Sheets[0].Name | Should -Be 'Secure'
            $summary.Schema.Worksheets[0].Name | Should -Be 'Secure'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $doc = Get-OfficeExcel -Path $path -Password 'secret'
        try {
            $doc.Sheets[0].Cell(1, 1, 'Saved with retained encryption', $null, $null)
            $saved = $doc | Save-OfficeExcel -PassThru -ErrorAction Stop
            $saved.FullName | Should -Be $path
            $saved | Should -BeOfType ([IO.FileInfo])
            { Get-ZipEntriesLocal -Path $path } | Should -Throw
            $doc.Sheets[0].Cell(1, 1, 'Do not save on close', $null, $null)
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $doc = Get-OfficeExcel -Path $path -Password 'secret' -ReadOnly
        try {
            $savedValue = $null
            $doc.Sheets[0].TryGetCellText(1, 1, [ref] $savedValue) | Should -BeTrue
            $savedValue | Should -Be 'Saved with retained encryption'
        } finally { Close-OfficeExcel -Document $doc }
        $sourceHash = (Get-FileHash -LiteralPath $path).Hash

        $plainCopyPath = Join-Path $TestDrive 'EncryptedExcelPlainCopy.xlsx'
        $doc = Get-OfficeExcel -Path $path -Password 'secret'
        try {
            $doc | Save-OfficeExcel -Path $plainCopyPath -Password $null
            $doc.Sheets[0].Cell(1, 1, 'Updated plain copy', $null, $null)
            $doc | Save-OfficeExcel
        } finally {
            Close-OfficeExcel -Document $doc
        }

        { Get-ZipEntriesLocal -Path $plainCopyPath } | Should -Not -Throw
        $plainCopy = Get-OfficeExcel -Path $plainCopyPath -ReadOnly
        try {
            $plainValue = $null
            $plainCopy.Sheets[0].TryGetCellText(1, 1, [ref] $plainValue) | Should -BeTrue
            $plainValue | Should -Be 'Updated plain copy'
        } finally {
            Close-OfficeExcel -Document $plainCopy
        }

        $encryptedCopyPath = Join-Path $TestDrive 'EncryptedExcelCopy.xlsx'
        $doc = Get-OfficeExcel -Path $path -Password 'secret'
        try {
            $doc | Save-OfficeExcel -Path $encryptedCopyPath -Password 'copy-secret'
            $doc.Sheets[0].Cell(1, 1, 'Updated encrypted copy', $null, $null)
            $doc | Save-OfficeExcel -ErrorAction Stop
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $encryptedCopy = Get-OfficeExcel -Path $encryptedCopyPath -Password 'copy-secret' -ReadOnly
        try {
            $copyValue = $null
            $encryptedCopy.Sheets[0].TryGetCellText(1, 1, [ref] $copyValue) | Should -BeTrue
            $copyValue | Should -Be 'Updated encrypted copy'
        } finally {
            Close-OfficeExcel -Document $encryptedCopy
        }

        (Get-FileHash -LiteralPath $path).Hash | Should -Be $sourceHash

        $doc = Get-OfficeExcel -Path $path -Password 'secret'
        try {
            $doc.Sheets[0].Cell(1, 1, 'Updated encrypted value', $null, $null)
            $doc | Save-OfficeExcel -Password 'secret'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $doc = Get-OfficeExcel -Path $path -Password 'secret' -ReadOnly
        try {
            $value = $null
            $doc.Sheets[0].TryGetCellText(1, 1, [ref] $value) | Should -BeTrue
            $value | Should -Be 'Updated encrypted value'
        } finally {
            Close-OfficeExcel -Document $doc
        }

    }

    It 'converts encrypted workbooks to HTML without changing or retaining the source' {
        $path = Join-Path $TestDrive 'EncryptedExcelHtml.xlsx'
        $htmlPath = Join-Path $TestDrive 'EncryptedExcelHtml.html'
        New-OfficeExcel -Path $path -Password 'secret' {
            Add-OfficeExcelSheet -Name 'Secure' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Encrypted HTML value'
            }
        }

        $original = [Convert]::ToBase64String([IO.File]::ReadAllBytes($path))

        ConvertTo-OfficeExcelHtml -Path $path -Password 'secret' -OutputPath $htmlPath

        Test-Path -LiteralPath $htmlPath | Should -BeTrue
        Get-Content -LiteralPath $htmlPath -Raw | Should -Match 'Encrypted HTML value'
        [Convert]::ToBase64String([IO.File]::ReadAllBytes($path)) | Should -Be $original
        $exclusive = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $exclusive.Dispose()
    }

}
