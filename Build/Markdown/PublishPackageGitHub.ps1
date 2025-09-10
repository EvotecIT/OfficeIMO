Import-Module PSPublishModule -Force -ErrorAction Stop

$GitHubAccessToken = Get-Content -Raw 'C:\Support\Important\GithubAPI.txt'

$publishGitHubReleaseAssetSplat = @{
    ProjectPath          = "$PSScriptRoot\..\..\OfficeIMO.Markdown"
    GitHubAccessToken    = $GitHubAccessToken
    GitHubUsername       = "EvotecIT"
    GitHubRepositoryName = "OfficeIMO"
    IsPreRelease         = $false
    IncludeProjectNameInTag = $true
}

Publish-GitHubReleaseAsset @publishGitHubReleaseAssetSplat
