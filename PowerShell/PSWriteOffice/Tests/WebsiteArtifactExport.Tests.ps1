BeforeAll {
    $exportScript = Join-Path $PSScriptRoot '../Build/Export-WebsiteArtifacts.ps1'
    $repositoryRoot = Join-Path $TestDrive 'source'
    $sourceRoot = Join-Path $repositoryRoot 'PowerShell/PSWriteOffice'
    $outputRoot = Join-Path $TestDrive 'artifacts'
    $buildRoot = Join-Path $TestDrive 'builds'
    $sourceInput = Join-Path $sourceRoot 'Sources/Example.cs'
    $sourceManifest = Join-Path $sourceRoot 'PSWriteOffice.psd1'
    $engineInput = Join-Path $repositoryRoot 'OfficeIMO.Word/Engine.cs'
    $engineFiles = @($engineInput, (Join-Path $repositoryRoot 'OfficeIMO/Engine.cs'))
    foreach ($file in $engineFiles) {
        $null = New-Item -ItemType Directory -Path (Split-Path $file)
        'public class Engine { }' | Set-Content $file
    }
    $nativeWorkingRoot = Join-Path $TestDrive 'native-working-directory'
    $null = New-Item -ItemType Directory -Path (Join-Path $nativeWorkingRoot 'OfficeIMO.Unrelated')
    $null = New-Item -ItemType Directory -Path (Split-Path $sourceInput), (Join-Path $sourceRoot 'Build')
    $manifestText = "@{ ModuleVersion = '3.0.8'; RootModule = 'PSWriteOffice.psm1'; Author = 'Łódź'; CmdletsToExport = @('Get-OfficeExample') }"
    $manifestEncoding = [Text.UTF8Encoding]::new($true)
    [IO.File]::WriteAllText($sourceManifest, $manifestText, $manifestEncoding)
    'public class Example { }' | Set-Content $sourceInput
    '# Source module' | Set-Content (Join-Path $sourceRoot 'PSWriteOffice.psm1')
    @'
param($RunMode, [bool]$SignModule, [bool]$IncludePackages, [switch]$SkipInstall, [switch]$PassThru, $StagingPath)
if ($IncludePackages) { throw 'Website API generation must not run a product release.' }
if ($RunMode -ne 'Build' -or $SignModule -or -not $SkipInstall -or -not $PassThru) { throw 'Unsafe fixture build options.' }
if (Test-Path -LiteralPath $StagingPath) { throw 'Reused fixture staging directory.' }
$null = New-Item -ItemType Directory -Path $StagingPath
throw 'Fresh unsigned build reached without installation.'
'@ | Set-Content (Join-Path $sourceRoot 'Build/Build-Module.ps1')
    & git -C $repositoryRoot init --quiet
    & git -C $repositoryRoot add -- .
    $emptyHooks = Join-Path $TestDrive 'empty-git-hooks'
    $null = New-Item -ItemType Directory -Path $emptyHooks
    & git -C $sourceRoot -c user.name='Test Fixture' -c user.email='tests@example.invalid' -c commit.gpgsign=false -c core.hooksPath=$emptyHooks commit --quiet -m 'Create source fixture'
    if ($LASTEXITCODE -ne 0) { throw 'Could not commit the isolated source fixture.' }
}

AfterAll {
    # Git marks loose objects read-only; let Pester remove only its own TestDrive.
    Get-ChildItem -LiteralPath (Join-Path $repositoryRoot '.git') -Recurse -Force -File |
        ForEach-Object { $_.IsReadOnly = $false }
}

Describe 'Website API source build qualification' {
    BeforeEach {
        foreach ($file in $engineFiles) {
            'public class Engine { }' | Set-Content $file
        }
        & git -C $repositoryRoot add -- OfficeIMO.Word/Engine.cs OfficeIMO/Engine.cs
        'public class Example { }' | Set-Content $sourceInput
        [IO.File]::WriteAllText($sourceManifest, $manifestText, $manifestEncoding)
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

    It 'reads UTF-8 manifests from Git independently of the console encoding' {
        $previousEncoding = [Console]::OutputEncoding
        try {
            [Console]::OutputEncoding = [Text.Encoding]::ASCII
            { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot -ArtifactsOnly } |
                Should -Throw '*Fresh unsigned build reached without installation*'
        } finally {
            [Console]::OutputEncoding = $previousEncoding
        }
    }

    It 'rejects source-local retained build or custom export output before building' {
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot (Join-Path $sourceRoot 'out/builds') } |
            Should -Throw '*BuildOutputRoot must be outside the source checkout*'
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot (Join-Path $sourceRoot 'out/website-api') -BuildOutputRoot $buildRoot } |
            Should -Throw '*ArtifactsRoot must be outside the source checkout*'
        Test-Path (Join-Path $sourceRoot 'out') | Should -BeFalse
    }

    It 'rejects changed source before running a build or importing a module' {
        'public class Changed { }' | Set-Content $sourceInput
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
            Should -Throw '*does not match the selected API and build inputs*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'rejects uncommitted <Package> source outside the module from an unrelated working directory' -ForEach @(
        @{ Package = 'OfficeIMO.Word' }
        @{ Package = 'OfficeIMO' }
    ) {
        'public class ChangedEngine { }' | Set-Content (Join-Path $repositoryRoot "$Package/Engine.cs")
        Push-Location $nativeWorkingRoot
        try {
            { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
                Should -Throw '*does not match the selected API and build inputs*'
            Test-Path $outputRoot | Should -BeFalse
        } finally { Pop-Location }
    }

    It 'rejects untracked <Package> source included by SDK compile globs from an unrelated working directory' -ForEach @(
        @{ Package = 'OfficeIMO.Word' }
        @{ Package = 'OfficeIMO' }
    ) {
        $untrackedEngine = Join-Path $repositoryRoot "$Package/UntrackedEngine.cs"
        Push-Location $nativeWorkingRoot
        try {
            'public class UntrackedEngine { }' | Set-Content $untrackedEngine
            { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
                Should -Throw '*does not match the selected API and build inputs*'
            Test-Path $outputRoot | Should -BeFalse
        } finally {
            Remove-Item -LiteralPath $untrackedEngine
            Pop-Location
        }
    }

    It 'rejects staged source changes even if the working file has been changed back' {
        'public class Changed { }' | Set-Content $sourceInput
        & git -C $sourceRoot add -- Sources/Example.cs
        'public class Engine { }' | Set-Content $engineInput
        & git -C $repositoryRoot add -- OfficeIMO.Word/Engine.cs
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
        [IO.File]::WriteAllText($sourceManifest, $manifestText.Replace("RootModule = 'PSWriteOffice.psm1'", "RootModule = 'Other.psm1'"), $manifestEncoding)
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
            Should -Throw '*source manifest does not match this checkout*'
        [IO.File]::WriteAllText($sourceManifest, $manifestText.Replace("'3.0.8'", "'3.0.9'"), $manifestEncoding)
        { & $exportScript -RepositoryRoot $sourceRoot -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
            Should -Throw '*Fresh unsigned build reached without installation*'
        Test-Path $outputRoot | Should -BeFalse
    }

    It 'requires the OfficeIMO module root rather than a copied or nested directory' {
        foreach ($root in @((Join-Path $TestDrive 'copied'), (Join-Path $sourceRoot 'nested'))) {
            $null = New-Item -ItemType Directory -Path $root
            Copy-Item $sourceManifest (Join-Path $root 'PSWriteOffice.psd1')
            { & $exportScript -RepositoryRoot $root -ArtifactsRoot $outputRoot -BuildOutputRoot $buildRoot } |
                Should -Throw '*OfficeIMO PowerShell module source root*'
        }
        Test-Path $outputRoot | Should -BeFalse
    }
}
