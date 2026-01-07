Import-Module PSPublishModule -Force -ErrorAction Stop

$GitHubAccessToken = Get-Content -Raw 'C:\Support\Important\GithubAPI.txt'
$publishGitHubReleaseAssetSplat = @{
    ProjectPath             = @(
        "$PSScriptRoot\..\OfficeIMO.CSV"
        "$PSScriptRoot\..\OfficeIMO.Excel"
        "$PSScriptRoot\..\OfficeIMO.Markdown"
        "$PSScriptRoot\..\OfficeIMO.Word"
    )
    GitHubAccessToken       = $GitHubAccessToken
    GitHubUsername          = "EvotecIT"
    GitHubRepositoryName    = "OfficeIMO"
    IsPreRelease            = $false
    IncludeProjectNameInTag = $true
}
Publish-GitHubReleaseAsset @publishGitHubReleaseAssetSplat
