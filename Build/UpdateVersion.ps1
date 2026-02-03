Import-Module PSPublishModule -Force -ErrorAction Stop

. "$PSScriptRoot\_ReleaseConfig.ps1"
$config = $OfficeIMOReleaseConfig

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")

Invoke-DotNetRepositoryRelease `
    -Path $repoRoot `
    -ExpectedVersion $config.ExpectedVersion `
    -ExpectedVersionMap $config.ExpectedVersionMap `
    -ExcludeProject $config.ExcludeProject `
    -NugetSource $config.NugetSource `
    -IncludePrerelease:$config.IncludePrerelease `
    -SkipPack
