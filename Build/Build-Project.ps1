param(
    [string] $ConfigPath = "$PSScriptRoot\project.build.json",
    [switch] $UpdateVersions,
    [switch] $Build,
    [switch] $PublishNuget,
    [switch] $PublishGitHub
)

Import-Module PSPublishModule -Force -ErrorAction Stop

$invokeParams = @{
    ConfigPath = $ConfigPath
}
if ($UpdateVersions) { $invokeParams.UpdateVersions = $true }
if ($Build) { $invokeParams.Build = $true }
if ($PublishNuget) { $invokeParams.PublishNuget = $false }
if ($PublishGitHub) { $invokeParams.PublishGitHub = $false }

Invoke-ProjectBuild @invokeParams
