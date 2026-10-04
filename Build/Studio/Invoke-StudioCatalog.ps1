param(
    [ValidateSet('Prepare', 'Submit', 'Status')][string] $Action = 'Prepare',
    [Parameter(Mandatory)][string] $OutputPath,
    [string] $AssetRoot,
    [string] $DeliveryReleaseId,
    [ValidateSet('winget', 'store', 'all')][string] $Channel = 'all',
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
if ($Execute -and $Action -ne 'Submit') { throw 'Execute is only valid for Submit.' }
$arguments = @('release', 'catalog', $Action.ToLowerInvariant(), '--config', (Join-Path $PSScriptRoot 'powerforge.catalog.json'), '--out', $OutputPath)
if ($Action -eq 'Prepare') {
    if (!$AssetRoot -or !$DeliveryReleaseId) { throw 'Prepare requires AssetRoot and DeliveryReleaseId.' }
    $arguments += @('--manifest', (Join-Path $AssetRoot 'windows-release-manifest.json'),
        '--checksums', (Join-Path $AssetRoot 'windows-SHA256SUMS.txt'), '--asset-root', $AssetRoot, '--delivery-release', $DeliveryReleaseId)
} elseif ($Action -eq 'Submit') {
    $arguments += @('--channel', $Channel)
    if ($Execute) { $arguments += '--execute' }
}
& powerforge @arguments
if ($LASTEXITCODE -ne 0) { throw "Studio catalog $Action failed. Inspect the catalog receipt before retrying." }
