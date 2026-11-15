#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Regenerates the website's officeimo command catalog from the tool itself.

.DESCRIPTION
    Builds OfficeIMO.Tool, runs `officeimo help --format json` and writes the
    result to data/cli/commands.json with a docsUrl for each command area.
    The tool's help text is the source of truth for commands, usage lines and
    options, so the website cannot list a command the tool does not dispatch.
    A command area without a documentation page fails the script.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Split-Path -Parent $PSScriptRoot),
    [string] $RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string] $Framework = 'net10.0',
    [string] $Configuration = 'Release',
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'

# The page that documents each command area. A new area must be added here.
$docsUrls = [ordered]@{
    'convert'     = '/docs/cli/convert/'
    'read'        = '/docs/cli/extract-and-inspect/'
    'inspect'     = '/docs/cli/extract-and-inspect/#inspect-a-file'
    'html'        = '/docs/cli/html/'
    'reader'      = '/docs/cli/extract-and-inspect/'
    'markup'      = '/docs/cli/markup/'
    'tabular'     = '/docs/cli/convert/#spreadsheets-and-delimited-files'
    'workflow'    = '/docs/cli/batch-and-jobs/'
    'invoice'     = '/docs/cli/invoices/'
    'pdf'         = '/docs/cli/pdf/'
    'provenance'  = '/docs/cli/before-sharing/#content-credentials-and-ai-labels'
    'hidden-data' = '/docs/cli/before-sharing/#hidden-data'
    'agent'       = '/docs/cli/mcp-and-agents/#agent-commands'
    'mcp'         = '/docs/cli/mcp-and-agents/'
}

if (-not $OutputPath) {
    $OutputPath = Join-Path $SiteRoot 'data/cli/commands.json'
}

$project = Join-Path $RepositoryRoot 'OfficeIMO.Tool/OfficeIMO.Tool.csproj'
if (-not (Test-Path -LiteralPath $project -PathType Leaf)) {
    throw "OfficeIMO.Tool project was not found: $project"
}

& dotnet build $project -f $Framework -c $Configuration -nologo -v quiet | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "Building OfficeIMO.Tool failed with exit code $LASTEXITCODE."
}

$toolPath = Join-Path $RepositoryRoot "OfficeIMO.Tool/bin/$Configuration/$Framework/OfficeIMO.Tool.dll"
$json = & dotnet $toolPath help --format json
if ($LASTEXITCODE -ne 0) {
    throw "officeimo help --format json failed with exit code $LASTEXITCODE."
}

$catalog = ($json -join [Environment]::NewLine) | ConvertFrom-Json
if ($catalog.schema -ne 'officeimo.tool.commands.v1') {
    throw "Unexpected command catalog schema '$($catalog.schema)'."
}

$missing = @($catalog.groups | Where-Object { -not $docsUrls.Contains($_.id) } | ForEach-Object id)
if ($missing.Count -gt 0) {
    throw "Add a docsUrl in Update-CliCatalog.ps1 for: $($missing -join ', ')."
}
$stale = @($docsUrls.Keys | Where-Object { $_ -notin $catalog.groups.id })
if ($stale.Count -gt 0) {
    throw "Remove docsUrl entries for command areas the tool no longer has: $($stale -join ', ')."
}

$groups = foreach ($group in $catalog.groups) {
    $projected = [ordered]@{}
    foreach ($property in $group.PSObject.Properties) {
        $projected[$property.Name] = $property.Value
        if ($property.Name -eq 'summary') {
            $projected['docsUrl'] = $docsUrls[$group.id]
        }
    }
    [pscustomobject] $projected
}

$site = [ordered]@{
    schema    = $catalog.schema
    tool      = $catalog.tool
    groups    = @($groups)
    exitCodes = $catalog.exitCodes
}

$parent = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $parent | Out-Null
}
[IO.File]::WriteAllText(
    $OutputPath,
    ($site | ConvertTo-Json -Depth 20) + [Environment]::NewLine,
    [Text.UTF8Encoding]::new($false))

$subcommandCount = ($catalog.groups | ForEach-Object { @($_.subcommands).Count } | Measure-Object -Sum).Sum
Write-Host "Wrote $OutputPath ($(@($catalog.groups).Count) areas, $subcommandCount commands, $(@($catalog.exitCodes).Count) exit codes)."
