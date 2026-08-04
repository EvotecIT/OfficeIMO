[CmdletBinding()]
param(
    [string] $ManifestPath = (Join-Path $PSScriptRoot '..\OfficeIMO.Rtf.Tests\Documents\RtfCorpus\corpus-manifest.json')
)

$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot 'ProducerCorpus/ExternalEvidenceVerifier/ExternalEvidenceVerifier.csproj'
& dotnet run --project $project --framework net8.0 -- rtf (Resolve-Path -LiteralPath $ManifestPath).Path
if ($LASTEXITCODE -ne 0) { throw 'External RTF producer evidence failed.' }
