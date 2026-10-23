<#
.SYNOPSIS
Exercises the npm path gate against an isolated synthetic Git history.
.DESCRIPTION
Uses local commits only, without npm, network calls, Pester or production files.
Creates and removes a uniquely named repository under the scratch root.
.PARAMETER ScratchRoot
Parent for task-owned scratch. Defaults to EVOTEC_SCRATCH_ROOT or D:/Dev/Scratch.
.EXAMPLE
./Build/BrowserExports/Test-NpmPackageChanges.ps1
.OUTPUTS
PSCustomObject reporting the number of passed cases and PowerShell version.
#>
[CmdletBinding()]
param(
    [string] $ScratchRoot = $(if ($env:EVOTEC_SCRATCH_ROOT) { $env:EVOTEC_SCRATCH_ROOT } else { 'D:/Dev/Scratch' })
)

$ErrorActionPreference = 'Stop'
$gate = Join-Path $PSScriptRoot 'Get-NpmPackageChanges.ps1'
$scratchParent = (Resolve-Path -LiteralPath $ScratchRoot).ProviderPath
$scratch = Join-Path $scratchParent ('officeimo-npm-path-gate-' + [guid]::NewGuid().ToString('N'))
$repository = Join-Path $scratch 'repository with spaces'
$passed = 0

function Invoke-TestGit {
    param([string[]] $GitArguments)
    $output = & git -C $repository @GitArguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Synthetic Git operation failed: $GitArguments`n$($output -join "`n")"
    }
    $output
}

function Set-TestFile {
    param([string] $Path, [string] $Content = 'synthetic npm input')
    $target = Join-Path $repository $Path
    $null = [IO.Directory]::CreateDirectory((Split-Path -Parent $target))
    [IO.File]::WriteAllText($target, $Content, [Text.UTF8Encoding]::new($false))
}

function Save-TestCommit {
    $null = Invoke-TestGit -GitArguments @('add', '--all')
    $null = Invoke-TestGit -GitArguments @('commit', '--quiet', '--no-verify', '--no-gpg-sign', '-m', 'Synthetic change')
    Invoke-TestGit -GitArguments @('rev-parse', 'HEAD')
}

function Assert-GateResult {
    param(
        [string] $Name,
        [string] $Base,
        [string] $Source = 'HEAD',
        [string[]] $ExpectedPaths = @()
    )
    $result = @(& $gate -BaseCommit $Base -SourceCommit $Source -RepositoryRoot $repository)
    if ($result.Count -ne 1 -or $result[0].Changed -isnot [bool] -or $result[0].Paths -isnot [array]) {
        throw "${Name}: expected one result with Boolean Changed and array Paths."
    }
    $actual = $result[0]
    if ($actual.Changed -ne ($ExpectedPaths.Count -gt 0) -or $actual.Paths.Count -ne $ExpectedPaths.Count) {
        throw "${Name}: incorrect change decision or path count."
    }
    foreach ($path in $ExpectedPaths) {
        if ($actual.Paths -cnotcontains $path) {
            throw "${Name}: a required release input is missing."
        }
    }
    $script:passed++
    Write-Host "PASS $Name"
}

function Assert-GateError {
    param([string] $Name, [string] $Base, [string] $Source = 'HEAD')
    $failed = $false
    try {
        $null = & $gate -BaseCommit $Base -SourceCommit $Source -RepositoryRoot $repository
    } catch {
        $failed = $true
    }
    if (-not $failed) {
        throw "${Name}: expected an error rather than a release decision."
    }
    $script:passed++
    Write-Host "PASS $Name"
}

$null = [IO.Directory]::CreateDirectory($repository)
try {
    $null = Invoke-TestGit -GitArguments @('-c', 'init.defaultBranch=main', 'init', '--quiet', '--template=')
    $settings = @{
        'user.name' = 'Synthetic npm gate test'
        'user.email' = 'npm-gate@example.invalid'
        'commit.gpgsign' = 'false'
        'tag.gpgSign' = 'false'
        'core.autocrlf' = 'false'
        'core.hooksPath' = (Join-Path $scratch 'no-hooks')
    }
    foreach ($setting in $settings.GetEnumerator()) {
        $null = Invoke-TestGit -GitArguments @('config', '--local', $setting.Key, $setting.Value)
    }
    Set-TestFile 'seed.txt'
    Set-TestFile 'OfficeIMO.JavaScript/src/delete.ts'
    Set-TestFile 'OfficeIMO.JavaScript/src/rename.ts'
    Set-TestFile 'OfficeIMO.JavaScript/test/move.ts'
    $baseline = Save-TestCommit
    $null = Invoke-TestGit -GitArguments @('tag', '-a', 'npm-published', '-m', 'Synthetic published source')
    Assert-GateResult -Name 'IdenticalCommits' -Base $baseline

    $releaseCases = @(
        @{ Name = 'JavaScriptSource'; Path = 'OfficeIMO.JavaScript/src/feature.ts' }
        @{ Name = 'Bundle'; Path = 'OfficeIMO.JavaScript/bundles/officeimo.js' }
        @{ Name = 'BuildScript'; Path = 'OfficeIMO.JavaScript/scripts/bundles.mjs' }
        @{ Name = 'PackageManifest'; Path = 'OfficeIMO.JavaScript/package.json' }
        @{ Name = 'PackageLock'; Path = 'OfficeIMO.JavaScript/package-lock.json' }
        @{ Name = 'CompilerConfig'; Path = 'OfficeIMO.JavaScript/tsconfig.json' }
        @{ Name = 'PackageLicense'; Path = 'OfficeIMO.JavaScript/LICENSE' }
        @{ Name = 'PackageReadme'; Path = 'OfficeIMO.JavaScript/README.md' }
        @{ Name = 'NpmWorkflow'; Path = '.github/workflows/npm-release.yml' }
        @{ Name = 'PathGate'; Path = 'Build/BrowserExports/Get-NpmPackageChanges.ps1' }
        @{ Name = 'AssetParityGate'; Path = 'Build/BrowserExports/Test-NpmAssetParity.ps1' }
        @{ Name = 'VersionGate'; Path = 'Build/BrowserExports/Get-NpmReleaseVersion.ps1' }
        @{ Name = 'ReleaseEligibilityGate'; Path = 'Build/BrowserExports/Get-NpmReleaseEligibility.ps1' }
    )
    foreach ($case in $releaseCases) {
        $null = Invoke-TestGit -GitArguments @('checkout', '--quiet', '-b', $case.Name, $baseline)
        Set-TestFile $case.Path
        $source = Save-TestCommit
        Assert-GateResult -Name $case.Name -Base 'npm-published' -Source $source -ExpectedPaths @($case.Path)
    }

    $null = Invoke-TestGit -GitArguments @('checkout', '--quiet', '-b', 'ExcludedInputs', $baseline)
    foreach ($path in @(
        'OfficeIMO.Word/Document.cs'
        'OfficeIMO.JavaScript.Assets/OfficeIMO.JavaScript.Assets.csproj'
        'Directory.Build.props'
        'OfficeIMO.JavaScript/test/sample.mjs'
        'OfficeIMO.JavaScript/test/consumer/tsconfig.json'
        'OfficeIMO.JavaScript/fixtures/sample.json'
        'OfficeIMO.JavaScript/src-other/sample.ts'
        'OfficeIMO.JavaScript/README.other.md'
        'Build/BrowserExports/Test-NpmPackageChanges.ps1'
        'Build/BrowserExports/Test-NpmConsumer.ps1'
        '.github/workflows/dotnet-tests.yml'
    )) {
        Set-TestFile $path
    }
    $null = Save-TestCommit
    Assert-GateResult -Name 'DotNetTestsFixturesAndNeighborPathsExcluded' -Base $baseline

    $null = Invoke-TestGit -GitArguments @('checkout', '--quiet', '-b', 'Deletion', $baseline)
    $null = Invoke-TestGit -GitArguments @('rm', '--quiet', 'OfficeIMO.JavaScript/src/delete.ts')
    $null = Save-TestCommit
    Assert-GateResult -Name 'DeletedSource' -Base $baseline -ExpectedPaths @('OfficeIMO.JavaScript/src/delete.ts')

    foreach ($case in @(
        @{ Name = 'RenameWithinSource'; Old = 'OfficeIMO.JavaScript/src/rename.ts'; New = 'OfficeIMO.JavaScript/src/new name.ts'; Expected = @('OfficeIMO.JavaScript/src/rename.ts', 'OfficeIMO.JavaScript/src/new name.ts') }
        @{ Name = 'RenameOutOfSource'; Old = 'OfficeIMO.JavaScript/src/rename.ts'; New = 'OfficeIMO.JavaScript/test/renamed.ts'; Expected = @('OfficeIMO.JavaScript/src/rename.ts') }
        @{ Name = 'RenameIntoSource'; Old = 'OfficeIMO.JavaScript/test/move.ts'; New = 'OfficeIMO.JavaScript/src/moved.ts'; Expected = @('OfficeIMO.JavaScript/src/moved.ts') }
    )) {
        $null = Invoke-TestGit -GitArguments @('checkout', '--quiet', '-b', $case.Name, $baseline)
        $null = Invoke-TestGit -GitArguments @('mv', $case.Old, $case.New)
        $null = Save-TestCommit
        Assert-GateResult -Name $case.Name -Base $baseline -ExpectedPaths $case.Expected
    }

    $null = Invoke-TestGit -GitArguments @('checkout', '--quiet', '-b', 'Cumulative', $baseline)
    Set-TestFile 'OfficeIMO.JavaScript/src/unpublished.ts'
    $javascriptCommit = Save-TestCommit
    $null = Invoke-TestGit -GitArguments @('tag', 'npm-reservation-only')
    $null = Invoke-TestGit -GitArguments @('checkout', '--quiet', '-b', 'DotNetSideBranch', $javascriptCommit)
    Set-TestFile 'OfficeIMO.Word/Document.cs'
    $dotnetCommit = Save-TestCommit
    $null = Invoke-TestGit -GitArguments @('checkout', '--quiet', 'Cumulative')
    $null = Invoke-TestGit -GitArguments @('merge', '--quiet', '--no-ff', '--no-verify', '--no-gpg-sign', '-m', 'Intermediate .NET merge', $dotnetCommit)
    Assert-GateResult -Name 'CumulativeSincePublishedAcrossDotNetMerge' -Base 'npm-published' -ExpectedPaths @('OfficeIMO.JavaScript/src/unpublished.ts')
    Assert-GateResult -Name 'IntermediateDotNetMergeAloneExcluded' -Base $javascriptCommit

    $defaultSource = & $gate -BaseCommit 'npm-published' -RepositoryRoot $repository
    if (-not $defaultSource.Changed -or $defaultSource.Paths.Count -ne 1 -or $defaultSource.Paths[0] -cne 'OfficeIMO.JavaScript/src/unpublished.ts') {
        throw 'DefaultSourceUsesHead: expected the current synthetic HEAD.'
    }
    $passed++
    Write-Host 'PASS DefaultSourceUsesHead'

    Assert-GateError -Name 'MissingBaselineRejected' -Base 'missing-published-tag'
    Assert-GateError -Name 'MissingSourceRejected' -Base $baseline -Source 'missing-source'
    Assert-GateError -Name 'OptionLikeRefRejected' -Base '--all'
    Assert-GateError -Name 'BlobRefRejected' -Base "${baseline}:seed.txt"
    Assert-GateError -Name 'TreeRefRejected' -Base "${baseline}^{tree}"
    Assert-GateError -Name 'NewerBaselineRejected' -Base 'HEAD' -Source $baseline
    Assert-GateError -Name 'DivergentHistoryRejected' -Base 'JavaScriptSource' -Source 'PackageReadme'

    # Git trees can contain newline names even on hosts that cannot create them
    # in a working directory. Exercise raw NUL parsing through the index.
    $blob = Invoke-TestGit -GitArguments @('rev-parse', "${baseline}:seed.txt")
    $unusualPaths = @(
        'OfficeIMO.JavaScript/src/space name.ts'
        ('OfficeIMO.JavaScript/src/' + [char]0x017C + '.ts')
        "OfficeIMO.JavaScript/src/line`nbreak.ts"
    )
    foreach ($path in $unusualPaths) {
        $null = Invoke-TestGit -GitArguments @('-c', 'core.protectNTFS=false', 'update-index', '--add', '--cacheinfo', "100644,$blob,$path")
    }
    $tree = Invoke-TestGit -GitArguments @('write-tree')
    $source = Invoke-TestGit -GitArguments @('commit-tree', $tree, '-p', 'HEAD', '-m', 'Synthetic unusual paths')
    Assert-GateResult -Name 'NulDelimitedSpacesUnicodeAndNewlines' -Base 'HEAD' -Source $source -ExpectedPaths $unusualPaths
} finally {
    # This unique leaf was created by this run and contains only synthetic data.
    # Verify containment and reject reparse points before removing it.
    $resolvedScratch = (Resolve-Path -LiteralPath $scratch).ProviderPath
    if ((Split-Path -Parent $resolvedScratch) -ne $scratchParent) {
        throw 'Scratch cleanup refused: unexpected parent.'
    }
    $entries = @(Get-Item -LiteralPath $resolvedScratch -Force) + @(Get-ChildItem -LiteralPath $resolvedScratch -Recurse -Force)
    if ($entries | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }) {
        throw "Scratch cleanup refused: reparse point under '$resolvedScratch'."
    }
    foreach ($entry in $entries) {
        $protectedAttributes = [IO.FileAttributes]::ReadOnly -bor [IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::System
        if ($entry.Attributes -band $protectedAttributes) {
            $entry.Attributes = $entry.Attributes -band (-bnot $protectedAttributes)
        }
    }
    Remove-Item -LiteralPath $resolvedScratch -Recurse -ErrorAction Stop
}

[pscustomobject]@{
    Passed = $passed
    PowerShell = $PSVersionTable.PSVersion.ToString()
    ScratchRemoved = -not (Test-Path -LiteralPath $scratch)
}
