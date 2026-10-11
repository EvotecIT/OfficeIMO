BeforeAll {
    $exportScript = Join-Path $PSScriptRoot '../Build/Export-WebsiteArtifacts.ps1'
    $sourceRoot = Join-Path $TestDrive 'source'
    $outputRoot = Join-Path $TestDrive 'artifacts'
    $buildRoot = Join-Path $TestDrive 'builds'
    $sourceInput = Join-Path $sourceRoot 'Sources/Example.cs'
    $sourceManifest = Join-Path $sourceRoot 'PSWriteOffice.psd1'
    $null = New-Item -ItemType Directory -Path (Split-Path $sourceInput), (Join-Path $sourceRoot 'Build')
    $manifestText = "@{ ModuleVersion = '3.0.8'; RootModule = 'PSWriteOffice.psm1'; CmdletsToExport = @('Get-OfficeExample') }"
    $manifestText | Set-Content $sourceManifest
    'public class Example { }' | Set-Content $sourceInput
    '# Source module' | Set-Content (Join-Path $sourceRoot 'PSWriteOffice.psm1')
    @'
param($RunMode, [bool]$SignModule, [switch]$SkipInstall, [switch]$PassThru, $StagingPath)
if ($RunMode -ne 'Build' -or $SignModule -or -not $SkipInstall -or -not $PassThru) { throw 'Unsafe fixture build options.' }
if (Test-Path -LiteralPath $StagingPath) { throw 'Reused fixture staging directory.' }
$null = New-Item -ItemType Directory -Path $StagingPath
throw 'Fresh unsigned build reached without installation.'
'@ | Set-Content (Join-Path $sourceRoot 'Build/Build-Module.ps1')
    & git -C $sourceRoot init --quiet
    & git -C $sourceRoot add -- .
    $emptyHooks = Join-Path $TestDrive 'empty-git-hooks'
    $null = New-Item -ItemType Directory -Path $emptyHooks
    & git -C $sourceRoot -c user.name='Test Fixture' -c user.email='tests@example.invalid' -c commit.gpgsign=false -c core.hooksPath=$emptyHooks commit --quiet -m 'Create source fixture'
    if ($LASTEXITCODE -ne 0) { throw 'Could not commit the isolated source fixture.' }
}

AfterAll {
    # Git marks loose objects read-only; let Pester remove only its own TestDrive.
    Get-ChildItem -LiteralPath (Join-Path $sourceRoot '.git') -Recurse -Force -File |
        ForEach-Object { $_.IsReadOnly = $false }
}

Describe 'Website API source build qualification' {
    BeforeEach {
        'public class Example { }' | Set-Content $sourceInput
        $manifestText | Set-Content $sourceManifest
        '# Source module' | Set-Content (Join-Path $sourceRoot 'PSWriteOffice.psm1')
        & git -C $sourceRoot add -- Sources/Example.cs PSWriteOffice.psd1 PSWriteOffice.psm1
    }

    It 'uses a new staging directory and an unsigned build without installation on every export' {
        1..2 | ForEach-Object {
            { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot -ArtifactsOnly } |
                Should -Throw '*Fresh unsigned build reached without installation*'
        }
        @(Get-ChildItem -LiteralPath $buildRoot -Directory).Count | Should -Be 2
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects changed source before running a build or importing a module' {
        'public class Changed { }' | Set-Content $sourceInput
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
            Should -Throw '*does not match the selected API and build inputs*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects staged source changes even if the working file has been changed back' {
        'public class Changed { }' | Set-Content $sourceInput
        & git -C $sourceRoot add -- Sources/Example.cs
        'public class Example { }' | Set-Content $sourceInput
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
            Should -Throw '*does not match the selected API and build inputs*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects uncommitted module code outside the compiled source tree' {
        '# Changed loader' | Set-Content (Join-Path $sourceRoot 'PSWriteOffice.psm1')
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
            Should -Throw '*does not match the selected API and build inputs*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects untracked build inputs' {
        $extra = Join-Path $sourceRoot 'Extra.ps1'
        try {
            '# Uncommitted module input' | Set-Content $extra
            { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
                Should -Throw '*does not match the selected API and build inputs*'
        } finally { Remove-Item -LiteralPath $extra }
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects changed manifest fields while allowing normal version resolution' {
        $manifestText.Replace("RootModule = 'PSWriteOffice.psm1'", "RootModule = 'Other.psm1'") | Set-Content $sourceManifest
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
            Should -Throw '*source manifest does not match this checkout*'
        $manifestText.Replace("'3.0.8'", "'3.0.9'") | Set-Content $sourceManifest
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
            Should -Throw '*Fresh unsigned build reached without installation*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'requires an exact Git source root rather than a copied or nested directory' {
        foreach ($root in @((Join-Path $TestDrive 'copied'), (Join-Path $sourceRoot 'nested'))) {
            $null = New-Item -ItemType Directory -Path $root
            Copy-Item $sourceManifest (Join-Path $root 'PSWriteOffice.psd1')
            { & $exportScript -RepositoryRoot $root -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
                Should -Throw '*exact Git source root*'
        }
        Test-Path $outputRoot | Should -BeFalse
    }
}
