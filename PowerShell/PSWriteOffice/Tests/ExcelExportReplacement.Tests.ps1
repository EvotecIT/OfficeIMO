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
            Export-OfficeExcel -Path $path -Sheet Data
        $originalHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash

        { [pscustomobject] @{} | Export-OfficeExcel -Path $path -ErrorAction Stop } |
            Should -Throw

        Test-Path -LiteralPath $path | Should -BeTrue
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -Be $originalHash
        $rows = @(Import-OfficeExcel -Path $path -Sheet Data)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'original'
        $rows[0].Amount | Should -Be 7
    }

    It 'replaces the existing workbook after a successful export' {
        $path = Join-Path $TestDrive 'SuccessfulReplacement.xlsx'
        [pscustomobject] @{ Name = 'original'; Amount = 7 } |
            Export-OfficeExcel -Path $path -Sheet Original

        [pscustomobject] @{ Name = 'replacement'; Amount = 19 } |
            Export-OfficeExcel -Path $path -Sheet Data -SaveOptions @{ ValidateOpenXml = $true }

        $rows = @(Import-OfficeExcel -Path $path -Sheet Data)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'replacement'
        $rows[0].Amount | Should -Be 19
        $saved = [OfficeIMO.Excel.ExcelDocument]::Load($path)
        try {
            $saved.Sheets.Count | Should -Be 1
            $saved.Sheets[0].Name | Should -Be 'Data'
        } finally { $saved.Dispose() }
    }

    It 'preserves the existing workbook when a locked destination prevents saving' -Skip:($env:OS -ne 'Windows_NT') {
        $path = Join-Path $TestDrive 'LockedReplacement.xlsx'
        [pscustomobject] @{ Name = 'original'; Amount = 7 } |
            Export-OfficeExcel -Path $path -Sheet Data
        $originalHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $lock = [System.IO.File]::Open($path, 'Open', 'Read', 'None')
        try {
            { [pscustomobject] @{ Name = 'replacement' } |
                Export-OfficeExcel -Path $path -ErrorAction Stop } | Should -Throw
        } finally {
            $lock.Dispose()
        }

        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -Be $originalHash
        $rows = @(Import-OfficeExcel -Path $path -Sheet Data)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'original'
    }
    It 'preserves omitted workbook metadata when appending rows' {
        $path = Join-Path $TestDrive 'AppendMetadata.xlsx'
        [pscustomobject] @{ Name = 'original'; Amount = 7 } |
            Export-OfficeExcel -Path $path -Sheet Data -DocumentProperty @{ Author = 'Finance'; Company = 'Contoso'; Title = 'Capacity' }
        [pscustomobject] @{ Name = 'appended'; Amount = 19 } |
            Export-OfficeExcel -Path $path -Sheet Data -Append

        $rows = @(Import-OfficeExcel -Path $path -Sheet Data)
        $rows.Count | Should -Be 2
        $saved = [OfficeIMO.Excel.ExcelDocument]::Load($path)
        try {
            $saved.Properties.Creator | Should -Be 'Finance'
            $saved.ApplicationProperties.Company | Should -Be 'Contoso'
            $saved.Properties.Title | Should -Be 'Capacity'
        } finally { $saved.Dispose() }
    }

}
