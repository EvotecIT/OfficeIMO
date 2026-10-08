param(
    [Parameter(Mandatory)][string] $OutputDirectory,
    [Parameter(Mandatory)][string] $ArtifactsDirectory
)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Independent creation qualification requires Windows DAO and Microsoft Access.' }
$root = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $root) { throw 'Use a fresh output directory. Existing user databases are never opened by this route.' }
$project = Join-Path $PSScriptRoot 'OfficeIMO.Access.Verification.csproj'
dotnet run --project $project -c Release --artifacts-path ([IO.Path]::GetFullPath($ArtifactsDirectory)) -- --create $root
if ($LASTEXITCODE -ne 0) { throw 'The public native creation example failed.' }
& (Join-Path $PSScriptRoot 'Test-PublicCreation.ps1') -CorpusDirectory $root
dotnet run --project $project -c Release --artifacts-path ([IO.Path]::GetFullPath($ArtifactsDirectory)) -- --verify-edits $root
if ($LASTEXITCODE -ne 0) { throw 'Decoding the Access-resaved output failed.' }
