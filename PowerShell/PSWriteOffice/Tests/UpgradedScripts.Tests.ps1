BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'Scripts written for PSWriteOffice 3.5' {
    It 'exports and imports a worksheet with -FilePath and -WorksheetName' {
        $path = Join-Path $TestDrive 'Legacy.xlsx'
        [pscustomobject] @{ Name = 'legacy'; Amount = 3 } |
            Export-OfficeExcel -FilePath $path -WorksheetName Data

        $rows = @(Import-OfficeExcel -FilePath $path -WorksheetName Data)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'legacy'
    }

    It 'Close-OfficeExcel <Outcome> when a workbook from New-OfficeExcel is closed <Case>' -ForEach @(
        @{ Case = 'without being saved'; Outcome = 'warns'; Save = $false; Warns = $true }
        @{ Case = 'after Save-OfficeExcel'; Outcome = 'stays silent'; Save = $true; Warns = $false }
    ) {
        $path = Join-Path $TestDrive "Live-$Save.xlsx"
        $workbook = New-OfficeExcel -Path $path
        if ($Save) { Save-OfficeExcel -Document $workbook }
        Close-OfficeExcel -Document $workbook -WarningVariable warnings -WarningAction SilentlyContinue

        if ($Warns) {
            @($warnings).Count | Should -Be 1
            "$($warnings[0])" | Should -BeLike "*'$path' was not written*Save-OfficeExcel*"
            Test-Path -LiteralPath $path | Should -BeFalse
        } else {
            @($warnings).Count | Should -Be 0
            Test-Path -LiteralPath $path | Should -BeTrue
        }
    }
}
