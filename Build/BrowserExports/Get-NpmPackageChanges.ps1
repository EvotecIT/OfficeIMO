<#
.SYNOPSIS
Finds npm release inputs changed since the last published npm source commit.
.DESCRIPTION
Compares the complete baseline-to-source tree diff, including deleted and renamed
paths. Both refs must resolve to commits and the baseline must be an ancestor of
the source. This read-only gate does not query npm or choose the published tag.
.PARAMETER BaseCommit
Commit or tag identifying the source of the last successfully published npm package.
Reservation tags and intermediate merges are not publication baselines.
.PARAMETER SourceCommit
Qualified source commit or tag to compare. Defaults to HEAD.
.PARAMETER RepositoryRoot
Repository to inspect. Defaults to the repository containing this script.
.EXAMPLE
./Build/BrowserExports/Get-NpmPackageChanges.ps1 -BaseCommit npm-v0.1.0 -SourceCommit HEAD
.OUTPUTS
PSCustomObject with Changed (Boolean) and Paths (String array of release inputs).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $BaseCommit,
    [ValidateNotNullOrEmpty()]
    [string] $SourceCommit = 'HEAD',
    [ValidateNotNullOrEmpty()]
    [string] $RepositoryRoot = "$PSScriptRoot/../.."
)

$ErrorActionPreference = 'Stop'
$repository = (Resolve-Path -LiteralPath $RepositoryRoot).ProviderPath

function Invoke-PathGateGit {
    # Capture raw UTF-8 output: PowerShell's native line pipeline splits filenames
    # containing newlines even when Git uses NUL delimiters.
    param([string[]] $GitArguments)

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = (Get-Command git -CommandType Application | Select-Object -First 1).Source
    $startInfo.WorkingDirectory = $repository
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [Text.Encoding]::UTF8
    $startInfo.StandardErrorEncoding = [Text.Encoding]::UTF8
    $startInfo.Arguments = ($GitArguments | ForEach-Object {
        '"' + (($_ -replace '(\\*)"', '$1$1\"') -replace '(\\+)$', '$1$1') + '"'
    }) -join ' '
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        $null = $process.Start()
        $output = $process.StandardOutput.ReadToEndAsync()
        $errorOutput = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        [pscustomobject]@{
            ExitCode = $process.ExitCode
            Output = $output.GetAwaiter().GetResult()
            Error = $errorOutput.GetAwaiter().GetResult()
        }
    } finally {
        $process.Dispose()
    }
}

$commits = foreach ($ref in @($BaseCommit, $SourceCommit)) {
    $resolved = Invoke-PathGateGit -GitArguments @('rev-parse', '--verify', '--end-of-options', "${ref}^{commit}")
    if ($resolved.ExitCode -ne 0) {
        throw "Cannot resolve '$ref' to a commit in '$repository': $($resolved.Error.Trim())"
    }
    $resolved.Output.Trim()
}
$ancestry = Invoke-PathGateGit -GitArguments @('merge-base', '--is-ancestor', $commits[0], $commits[1])
if ($ancestry.ExitCode -ne 0) {
    throw "BaseCommit must be an ancestor of SourceCommit; cannot compare '$BaseCommit' to '$SourceCommit'. $($ancestry.Error.Trim())"
}
$diff = Invoke-PathGateGit -GitArguments @('diff', '--name-only', '--no-renames', '-z', "$($commits[0])..$($commits[1])", '--')
if ($diff.ExitCode -ne 0) {
    throw "Cannot compare npm release inputs: $($diff.Error.Trim())"
}

$exactPaths = @(
    'OfficeIMO.JavaScript/package.json'
    'OfficeIMO.JavaScript/package-lock.json'
    'OfficeIMO.JavaScript/tsconfig.json'
    'OfficeIMO.JavaScript/LICENSE'
    'OfficeIMO.JavaScript/README.md'
    '.github/workflows/npm-release.yml'
    'Build/BrowserExports/Get-NpmPackageChanges.ps1'
    'Build/BrowserExports/Test-NpmAssetParity.ps1'
    'Build/BrowserExports/Get-NpmReleaseVersion.ps1'
    'Build/BrowserExports/Get-NpmReleaseEligibility.ps1'
)
$paths = @($diff.Output.Split([char]0) | Where-Object {
    $_ -cmatch '^OfficeIMO\.JavaScript/(src|bundles|scripts)/' -or $exactPaths -ccontains $_
})
[pscustomobject]@{
    Changed = $paths.Count -gt 0
    Paths = [string[]] $paths
}
