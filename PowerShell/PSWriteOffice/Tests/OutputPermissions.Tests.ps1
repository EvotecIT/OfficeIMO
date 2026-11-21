BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

# Files saved to a shared folder must stay readable by everyone that folder grants access to.
Describe 'Saved document permissions' -Skip:(-not $IsWindows -and $PSVersionTable.PSEdition -eq 'Core') {
    BeforeAll {
        $folder = Join-Path $TestDrive 'Shared'
        New-Item -ItemType Directory -Path $folder | Out-Null
        $networkService = [System.Security.Principal.SecurityIdentifier]::new('S-1-5-20')
        $acl = Get-Acl -LiteralPath $folder
        $acl.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new(
            $networkService, 'Read', 'ContainerInherit, ObjectInherit', 'None', 'Allow'))
        Set-Acl -LiteralPath $folder -AclObject $acl
        $rows = 1..2 | ForEach-Object { [pscustomobject] @{ Id = $_; Name = "Row $_" } }
    }

    It '<Name> gives a new file the destination folder''s inherited permissions' -ForEach @(
        @{ Name = 'Export-OfficeExcel'; File = 'Export.xlsx'; Save = { param($Path) $rows | Export-OfficeExcel -Path $Path } }
        @{ Name = 'New-OfficeWord'; File = 'New.docx'; Save = { param($Path) New-OfficeWord -Path $Path { WordParagraph -Text 'Shared' } } }
    ) {
        $path = Join-Path $folder $File
        & $Save $path

        $saved = Get-Acl -LiteralPath $path
        $saved.AreAccessRulesProtected | Should -BeFalse
        $inherited = @($saved.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]) |
            Where-Object { $_.IsInherited -and $_.IdentityReference -eq $networkService })
        $inherited.Count | Should -BeGreaterThan 0
    }
}
