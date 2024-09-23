param (
    $SolutionRoot = "$PSScriptRoot\.."
)

$NugetAPI = Get-Content -Raw -LiteralPath "C:\Support\Important\NugetOrg.txt"
#$GitHubAPI = Get-Content -Raw -LiteralPath "C:\Support\Important\GithubAPI.txt"

$ReleasePath = [io.path]::Combine($SolutionRoot, "OfficeIMO.Word", "bin", "Release")
$File = Get-ChildItem -Path $ReleasePath -Recurse -Filter "*.nupkg"

# publish to nuget.org
if ($File.Count -eq 1) {
    dotnet nuget push $File.FullName --api-key $NugetAPI --source https://api.nuget.org/v3/index.json

    #dotnet nuget add source --username evotecit --password $GitHubAPI --store-password-in-clear-text --name github "https://nuget.pkg.github.com/OWNER/index.json"
    #dotnet nuget push $File.FullName --api-key $GitHubAPI --source "github"
}