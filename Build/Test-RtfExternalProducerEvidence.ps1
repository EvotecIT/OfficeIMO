[CmdletBinding()]
param(
    [string] $ManifestPath = (Join-Path $PSScriptRoot '..\OfficeIMO.Rtf.Tests\Documents\RtfCorpus\corpus-manifest.json'),
    [string] $Configuration = 'Debug',
    [string] $Framework = 'net8.0',
    [switch] $NoRestore,
    [switch] $NoBuild
)

$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot 'ProducerCorpus/ExternalEvidenceVerifier/ExternalEvidenceVerifier.csproj'
$arguments = @('run', '--project', $project, '--configuration', $Configuration, '--framework', $Framework)
if ($NoRestore) { $arguments += '--no-restore' }
if ($NoBuild) { $arguments += '--no-build' }
$arguments += @('--', 'rtf', (Resolve-Path -LiteralPath $ManifestPath).Path)
& dotnet @arguments
if ($LASTEXITCODE -ne 0) { throw 'External RTF producer evidence failed.' }
