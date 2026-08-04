[CmdletBinding()]
param(
    [string] $ManifestPath = (Join-Path $PSScriptRoot '..\OfficeIMO.OpenDocument.Tests\Fixtures\producer-manifest.json')
)

$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot 'ProducerCorpus/ExternalEvidenceVerifier/ExternalEvidenceVerifier.csproj'
& dotnet run --project $project --framework net8.0 -- odf (Resolve-Path -LiteralPath $ManifestPath).Path
if ($LASTEXITCODE -ne 0) { throw 'External OpenDocument producer evidence failed.' }
