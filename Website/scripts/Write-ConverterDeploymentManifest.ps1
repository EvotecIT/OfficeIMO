[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $ConverterRoot,

    [string] $DeploymentId,

    [string] $DeploymentIdEnvironment = 'GITHUB_SHA'
)

$ErrorActionPreference = 'Stop'
$root = [System.IO.Path]::GetFullPath($ConverterRoot)
if (-not (Test-Path -LiteralPath $root -PathType Container)) {
    throw "Converter root '$root' does not exist."
}

if ([string]::IsNullOrWhiteSpace($DeploymentId) -and -not [string]::IsNullOrWhiteSpace($DeploymentIdEnvironment)) {
    $DeploymentId = [Environment]::GetEnvironmentVariable($DeploymentIdEnvironment)
}
if ([string]::IsNullOrWhiteSpace($DeploymentId)) {
    $DeploymentId = (& git rev-parse HEAD 2>$null).Trim()
}
if ($DeploymentId -notmatch '^[A-Fa-f0-9]{40}$|^[A-Fa-f0-9]{64}$') {
    throw "Converter deployment id '$DeploymentId' is not a source commit digest."
}

$manifestName = 'deployment-assets.json'
$assets = foreach ($file in Get-ChildItem -LiteralPath $root -Recurse -File | Sort-Object FullName) {
    $relativePath = [System.IO.Path]::GetRelativePath($root, $file.FullName).Replace('\', '/')
    if ($relativePath -eq $manifestName -or $relativePath.EndsWith('.br', [StringComparison]::OrdinalIgnoreCase) -or $relativePath.EndsWith('.gz', [StringComparison]::OrdinalIgnoreCase)) {
        continue
    }
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    [ordered]@{
        path = $relativePath
        bytes = $bytes.LongLength
        sha256 = [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($bytes)).ToLowerInvariant()
    }
}

if (-not ($assets | Where-Object path -EQ 'index.html')) {
    throw "Converter deployment '$root' does not contain index.html."
}

$manifest = [ordered]@{
    schemaVersion = 1
    deploymentId = $DeploymentId.ToLowerInvariant()
    assets = @($assets)
}
$manifestPath = Join-Path $root $manifestName
$manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding utf8NoBOM
Write-Output "Converter deployment manifest written: $($manifest.assets.Count) public assets for $($manifest.deploymentId)."
