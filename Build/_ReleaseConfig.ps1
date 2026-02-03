# Shared configuration for OfficeIMO build/release scripts.
$OfficeIMOReleaseConfig = @{
    ExpectedVersion           = $null
    ExpectedVersionMap        = @{
        "OfficeIMO.CSV"      = "0.1.X"
        "OfficeIMO.Excel"    = "0.6.X"
        "OfficeIMO.Markdown" = "0.5.X"
        "OfficeIMO.Word"     = "1.0.X"
    }
    ExcludeProject             = @("OfficeIMO.Visio", "OfficeIMO.Project")
    NugetSource                = @()
    IncludePrerelease          = $false
    OutputPath                 = $null
    PublishSource              = "https://api.nuget.org/v3/index.json"
    PublishApiKeyFilePath      = "C:\Support\Important\NugetOrgEvotec.txt"
    GitHubAccessTokenFilePath  = "C:\Support\Important\GithubAPI.txt"
    GitHubUsername             = "EvotecIT"
    GitHubRepositoryName       = "OfficeIMO"
    GitHubIsPreRelease          = $false
    GitHubIncludeProjectNameInTag = $true
}
