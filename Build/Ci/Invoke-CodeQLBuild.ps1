<#
.SYNOPSIS
Builds the complete production graph for traced CodeQL extraction.
.DESCRIPTION
Selects net10 for multi-targeted projects without imposing a global framework on
single-targeted solution roots. Keeps compiler analyzers and source generators.
.PARAMETER OriginalFrameworkSelection
Uses the original solution-wide net10 property for paired diagnostic builds.
.EXAMPLE
./Build/Ci/Invoke-CodeQLBuild.ps1
#>
param([switch] $OriginalFrameworkSelection)
$graphArguments = @()
$mainFramework = @()
if ($OriginalFrameworkSelection) {
    $mainFramework += @('--framework', 'net10.0')
} else {
    $graphArguments += '-p:OfficeIMOCodeQLBuild=true'
}
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
# Production projects live at the repository root. Exclude qualification
# hosts and the iOS workload; the portable source generator is built separately.
$projects = Get-ChildItem -Directory -Filter 'OfficeIMO.*' | Where-Object {
  $_.Name -notmatch '\.(Tests|VerifyTests|Benchmarks|Examples|AotSmoke|[^.]*Verification)(\.|$)' -and
  $_.Name -notin @('OfficeIMO.Studio.iOS', 'OfficeIMO.Data.Generators')
} | ForEach-Object { Get-ChildItem -LiteralPath $_.FullName -Filter *.csproj } | Sort-Object FullName
if (!$projects) { throw 'No production projects found for CodeQL.' }

$codeqlSolution = Join-Path $env:RUNNER_TEMP 'OfficeIMO.CodeQL.sln'
$artifacts = Join-Path $env:RUNNER_TEMP 'OfficeIMO.CodeQL.Artifacts'
dotnet new sln --name OfficeIMO.CodeQL --output $env:RUNNER_TEMP --format sln
dotnet sln $codeqlSolution add @($projects.FullName)

# CodeQL needs a representative production build, not every target framework.
# Every selected production project supports net10; this also captures Studio/AI
# without forcing their source into an unsupported net8 target.
# Keep the normal test/build workflow responsible for the full TFM matrix.
# Keep the original full restore: legacy targets still need their packages.
# Framework selection applies to compilation, after assets are available.
dotnet restore $codeqlSolution --artifacts-path $artifacts "-bl:$env:RUNNER_TEMP/codeql-1.binlog"
dotnet build $codeqlSolution --configuration Release @mainFramework --no-restore --artifacts-path $artifacts @graphArguments "-bl:$env:RUNNER_TEMP/codeql-2.binlog"

dotnet restore OfficeIMO.Data.Generators/OfficeIMO.Data.Generators.csproj --artifacts-path $artifacts "-bl:$env:RUNNER_TEMP/codeql-3.binlog"
dotnet build OfficeIMO.Data.Generators/OfficeIMO.Data.Generators.csproj --configuration Release --framework netstandard2.0 --no-restore --artifacts-path $artifacts @graphArguments "-bl:$env:RUNNER_TEMP/codeql-4.binlog"

dotnet restore Examples/OfficeIMO.AI.Example/OfficeIMO.AI.Example.csproj --artifacts-path $artifacts "-bl:$env:RUNNER_TEMP/codeql-5.binlog"
dotnet build Examples/OfficeIMO.AI.Example/OfficeIMO.AI.Example.csproj --configuration Release --framework net10.0 --no-restore --artifacts-path $artifacts @graphArguments "-bl:$env:RUNNER_TEMP/codeql-6.binlog"
