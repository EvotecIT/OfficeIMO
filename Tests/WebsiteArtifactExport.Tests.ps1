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
    $sourceInput = Join-Path $sourceRoot 'Sources/Example.cs'
    $null = New-Item -ItemType Directory -Path (Split-Path -Parent $sourceInput)
    'public class Example { }' | Set-Content $sourceInput
    & git -C $sourceRoot init --quiet
    & git -C $sourceRoot add -- PSWriteOffice.psd1 Sources/Example.cs
    $emptyHooks = Join-Path $TestDrive 'empty-git-hooks'
    $null = New-Item -ItemType Directory -Path $emptyHooks
    & git -C $sourceRoot -c user.name='Test Fixture' -c user.email='tests@example.invalid' -c commit.gpgsign=false -c core.hooksPath=$emptyHooks commit --quiet -m 'Create source fixture'
    if ($LASTEXITCODE -ne 0) { throw 'Could not commit the isolated source fixture.' }
    $copiedSource = Join-Path $TestDrive 'copied-source'
    $nestedSource = Join-Path $sourceRoot 'nested-copy'
    $null = New-Item -ItemType Directory -Path $copiedSource, $nestedSource
    Copy-Item $sourceManifest (Join-Path $copiedSource 'PSWriteOffice.psd1')
    Copy-Item $sourceManifest (Join-Path $nestedSource 'PSWriteOffice.psd1')
}

AfterAll {
    # Git marks loose objects read-only; let Pester remove only its own TestDrive.
    Get-ChildItem -LiteralPath (Join-Path $sourceRoot '.git') -Recurse -Force -File |
        ForEach-Object { $_.IsReadOnly = $false }
}

Describe 'Explicit website API bundle qualification' {
    BeforeEach {
        'public class Example { }' | Set-Content $sourceInput
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

    It 'rejects a copied source directory without a Git revision' {
        { & $exportScript -RepositoryRoot $copiedSource -ArtifactsRoot $outputRoot -SkipBuild -BuiltModuleManifestPath $bundleManifest } |
            Should -Throw '*exact Git source root*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects a nested copy instead of recording the enclosing repository commit' {
        { & $exportScript -RepositoryRoot $nestedSource -ArtifactsRoot $outputRoot -SkipBuild -BuiltModuleManifestPath $bundleManifest } |
            Should -Throw '*exact Git source root*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'resolves a relative bundle path from the PowerShell location' {
        Push-Location $bundleRoot
        try {
            { & $exportScript -RepositoryRoot $copiedSource -ArtifactsRoot $outputRoot -SkipBuild -BuiltModuleManifestPath './PSWriteOffice.psd1' } |
                Should -Throw '*exact Git source root*'
        } finally { Pop-Location }
        Test-Path $outputRoot | Should -BeFalse
    }
}
