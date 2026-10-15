#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Verifies the AI agent integrations page against the MCP server, the plugin and Docs/ROADMAP.md.

.DESCRIPTION
    data/agent_integrations.json is the public record of how agents use OfficeIMO. Its tools must be exactly the MCP
    tools OfficeIMO.Tool registers, its skills exactly the plugin's skills, its version the plugin's version, and the
    install configurations it publishes must launch the same server as the plugin's mcp.json. Planned listings point to
    an existing section of Docs/ROADMAP.md, which remains the single backlog.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Split-Path -Parent $PSScriptRoot),
    [string] $RepositoryRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = 'Stop'
$repository = (Resolve-Path -LiteralPath $RepositoryRoot).Path
$data = Get-Content -LiteralPath (Join-Path (Resolve-Path -LiteralPath $SiteRoot).Path 'data/agent_integrations.json') -Raw | ConvertFrom-Json
$pluginRoot = Join-Path $repository '.agents/plugins/officeimo-document-tools'
$plugin = Get-Content -LiteralPath (Join-Path $pluginRoot 'plugin.json') -Raw | ConvertFrom-Json
$pluginServer = (Get-Content -LiteralPath (Join-Path $pluginRoot 'mcp.json') -Raw | ConvertFrom-Json).mcpServers.officeimo
$problems = [System.Collections.Generic.List[string]]::new()

if ($data.pluginVersion -ne $plugin.version) {
    $problems.Add("pluginVersion is $($data.pluginVersion); the plugin is $($plugin.version).")
}

# Tools: exactly the names the server registers with [McpServerTool(Name = "...")].
$registered = Get-ChildItem -LiteralPath (Join-Path $repository 'OfficeIMO.Tool/Mcp') -Filter '*.cs' |
    ForEach-Object { [regex]::Matches((Get-Content -LiteralPath $_.FullName -Raw), '\[McpServerTool\([^\]]*?Name\s*=\s*"([^"]+)"') } |
    ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
$listed = @($data.tools.name)
foreach ($name in $registered) { if ($name -notin $listed) { $problems.Add("The server registers '$name', which the page doesn't list.") } }
foreach ($name in $listed) { if ($name -notin $registered) { $problems.Add("The page lists '$name', which the server doesn't register.") } }
if ($listed.Count -ne ($listed | Sort-Object -Unique).Count) { $problems.Add('A tool is listed twice.') }

# Skills: exactly the plugin's skill folders.
$skills = Get-ChildItem -LiteralPath (Join-Path $pluginRoot 'skills') -Directory | ForEach-Object Name | Sort-Object
$listedSkills = @($data.skills.name) | Sort-Object
if (($skills -join ',') -ne ($listedSkills -join ',')) {
    $problems.Add("The page lists skills '$($listedSkills -join ', ')'; the plugin has '$($skills -join ', ')'.")
}

# The install links must start the plugin's server, at the page's version.
$expectedArguments = @($pluginServer.args | ForEach-Object { $_ -replace '(?<=OfficeIMO\.Tool@).+$', $data.pluginVersion })
foreach ($property in 'mcpConfig', 'cursorConfig') {
    $config = ($data.$property -replace '\{version\}', $data.pluginVersion) | ConvertFrom-Json
    if ($config.command -ne $pluginServer.command -or (@($config.args) -join ' ') -ne ($expectedArguments -join ' ')) {
        $problems.Add("$property doesn't launch the plugin's server: $($config.command) $(@($config.args) -join ' ').")
    }
    if (-not $config.env.OFFICEIMO_MCP_ALLOWED_ROOTS) { $problems.Add("$property sets no OFFICEIMO_MCP_ALLOWED_ROOTS, so the server would refuse to start.") }
}
foreach ($client in $data.clients) {
    foreach ($command in @($client.commands)) {
        if ($command -match 'OfficeIMO\.Tool@(?!\{version\})') { $problems.Add("Client '$($client.id)' pins a version; use {version}.") }
    }
}

# States and roadmap links.
$anchors = [System.Collections.Generic.HashSet[string]]::new()
foreach ($line in Get-Content -LiteralPath (Join-Path $repository 'Docs/ROADMAP.md')) {
    if ($line -match '^#{1,6}\s+(.+?)\s*$') {
        [void] $anchors.Add((($Matches[1].ToLowerInvariant() -replace '[^\p{L}\p{Nd}\s_-]', '') -replace '\s', '-'))
    }
}
$stateIds = @($data.states.id)
foreach ($entry in @($data.clients) + @($data.directories)) {
    if ($entry.state -notin $stateIds) { $problems.Add("'$($entry.id)' uses unknown state '$($entry.state)'.") }
    if ($entry.state -eq 'planned') {
        if (-not $entry.roadmap) { $problems.Add("'$($entry.id)' is planned without a roadmap section.") }
        elseif (-not $anchors.Contains([string] $entry.roadmap)) { $problems.Add("'$($entry.id)' links to '#$($entry.roadmap)', which Docs/ROADMAP.md does not have.") }
    }
    if ($entry.state -eq 'listed' -and -not $entry.url) { $problems.Add("'$($entry.id)' is listed without a link to the listing.") }
}

if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Host "Agent integrations: $_" }
    throw "The agent integrations page has $($problems.Count) problem(s)."
}
Write-Host "Agent integrations: $($listed.Count) tools, $($skills.Count) skills and $(@($data.clients).Count) clients verified at $($data.pluginVersion)."
