Import-Module PSPublishModule -Force -ErrorAction Stop

. "$PSScriptRoot\_ReleaseConfig.ps1"
$config = $OfficeIMOReleaseConfig

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")

if (-not $config.ExpectedVersionMap -or $config.ExpectedVersionMap.Count -eq 0) {
    throw "ExpectedVersionMap is required to resolve project list for GitHub release."
}

$projectNames = @($config.ExpectedVersionMap.Keys)
if ($config.ExcludeProject) {
    $projectNames = $projectNames | Where-Object { $config.ExcludeProject -notcontains $_ }
}

$projectPaths = $projectNames | ForEach-Object { Join-Path $repoRoot $_ }

$GitHubAccessToken = Get-Content -Raw $config.GitHubAccessTokenFilePath
$publishGitHubReleaseAssetSplat = @{
    ProjectPath             = $projectPaths
    GitHubAccessToken       = $GitHubAccessToken
    GitHubUsername          = $config.GitHubUsername
    GitHubRepositoryName    = $config.GitHubRepositoryName
    IsPreRelease            = [bool]$config.GitHubIsPreRelease
    IncludeProjectNameInTag = [bool]$config.GitHubIncludeProjectNameInTag
}
Publish-GitHubReleaseAsset @publishGitHubReleaseAssetSplat
