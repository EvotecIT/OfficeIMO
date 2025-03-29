param (
    $SolutionRoot = "$PSScriptRoot\.."
)

$GitHubAccessToken = Get-Content -Raw 'C:\Support\Important\GithubAPI.txt'
$UserName = 'EvotecIT'
$GitHubRepositoryName = 'OfficeIMO'

$SolutionPath = [io.path]::Combine($SolutionRoot, 'OfficeImo.sln')
if (-not $SolutionRoot -or -not (Test-Path -Path $SolutionPath)) {
    Write-Host -Object "Solution not found at $SolutionPath" -ForegroundColor Red
    return
}

$ProjectPath = [io.path]::Combine($SolutionRoot, "OfficeIMO.Word", "OfficeIMO.Word.csproj")

[xml] $XML = Get-Content -Raw $ProjectPath

$Version = $XML.Project.PropertyGroup[0].VersionPrefix

$ZipPath = [io.path]::Combine($SolutionRoot, "OfficeIMO.Word", "bin", "Release", "OfficeIMO.Word.$Version.zip")
$IsPreRelease = $false
$TagName = "v$Version"

if (Test-Path -LiteralPath $ZipPath) {
    $sendGitHubReleaseSplat = @{
        GitHubUsername       = $UserName
        GitHubRepositoryName = $GitHubRepositoryName
        GitHubAccessToken    = $GitHubAccessToken
        TagName              = $TagName
        AssetFilePaths       = $ZipPath
        IsPreRelease         = $IsPreRelease
    }
    $StatusGithub = Send-GitHubRelease @sendGitHubReleaseSplat
    $StatusGithub
}