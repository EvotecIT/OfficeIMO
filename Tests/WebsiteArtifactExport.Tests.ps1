BeforeAll {
    $exportScript = Join-Path $PSScriptRoot '../Build/Export-WebsiteArtifacts.ps1'
    $sourceRoot = Join-Path $TestDrive 'source'
    $bundleRoot = Join-Path $TestDrive 'bundle'
    $null = New-Item -ItemType Directory -Path $sourceRoot, (Join-Path $bundleRoot 'en-US')
    $sourceManifest = Join-Path $sourceRoot 'PSWriteOffice.psd1'
    $bundleManifest = Join-Path $bundleRoot 'PSWriteOffice.psd1'
    $bundleHelp = Join-Path $bundleRoot 'en-US/PSWriteOffice-help.xml'
    $outputRoot = Join-Path $TestDrive 'artifacts'
    "@{ ModuleVersion = '3.0.8'; CmdletsToExport = @('Get-OfficeExample') }" | Set-Content $sourceManifest
}

Describe 'Explicit website API bundle qualification' {
    BeforeEach {
        "@{ ModuleVersion = '3.0.8' }" | Set-Content $bundleManifest
        '<helpItems><command><details><name>Get-OfficeExample</name></details></command></helpItems>' | Set-Content $bundleHelp
    }

    It 'requires an explicit decision to skip the build before using a package bundle' {
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuiltModuleManifestPath $bundleManifest } |
            Should -Throw '*requires -SkipBuild*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects stale bundles even when the exported command names are unchanged' {
        "@{ ModuleVersion = '3.0.6' }" | Set-Content $bundleManifest
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -SkipBuild -BuiltModuleManifestPath $bundleManifest } |
            Should -Throw '*complete bundle for PSWriteOffice 3.0.8*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects help for a different command before importing or writing output' {
        '<helpItems><command><details><name>Get-OfficeOther</name></details></command></helpItems>' | Set-Content $bundleHelp
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -SkipBuild -BuiltModuleManifestPath $bundleManifest } |
            Should -Throw '*help for every exported command*'
        Test-Path $outputRoot | Should -BeFalse
    }
}
