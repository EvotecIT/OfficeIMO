BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Excel workbook replacement' {
    It 'preserves the existing workbook when replacement input has no columns' {
        $path = Join-Path $TestDrive 'InvalidReplacement.xlsx'
        [pscustomobject] @{ Name = 'original'; Amount = 7 } |
            Export-OfficeExcel -Path $path -WorksheetName Data
        $originalHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash

        { [pscustomobject] @{} | Export-OfficeExcel -Path $path -ErrorAction Stop } |
            Should -Throw

        Test-Path -LiteralPath $path | Should -BeTrue
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -Be $originalHash
        $rows = @(Import-OfficeExcel -Path $path -WorksheetName Data)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'original'
        $rows[0].Amount | Should -Be 7
    }

    It 'replaces the existing workbook after a successful export' {
        $path = Join-Path $TestDrive 'SuccessfulReplacement.xlsx'
        [pscustomobject] @{ Name = 'original'; Amount = 7 } |
            Export-OfficeExcel -Path $path -WorksheetName Original

        [pscustomobject] @{ Name = 'replacement'; Amount = 19 } |
            Export-OfficeExcel -Path $path -WorksheetName Data -ValidateOpenXml

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName Data)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'replacement'
        $rows[0].Amount | Should -Be 19
        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        @($summary.Sheets).Count | Should -Be 1
        $summary.Sheets[0].Name | Should -Be 'Data'
    }

    It 'preserves the existing workbook when a locked destination prevents saving' -Skip:($env:OS -ne 'Windows_NT') {
        $path = Join-Path $TestDrive 'LockedReplacement.xlsx'
        [pscustomobject] @{ Name = 'original'; Amount = 7 } |
            Export-OfficeExcel -Path $path -WorksheetName Data
        $originalHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $lock = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
        try {
            { [pscustomobject] @{ Name = 'replacement' } |
                Export-OfficeExcel -Path $path -ErrorAction Stop } | Should -Throw
        } finally {
            $lock.Dispose()
        }

        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -Be $originalHash
        $rows = @(Import-OfficeExcel -Path $path -WorksheetName Data)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'original'
    }
}
