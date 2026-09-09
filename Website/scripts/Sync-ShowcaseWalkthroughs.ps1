[CmdletBinding()]
param([switch] $Check)

$ErrorActionPreference = 'Stop'
$websiteRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $websiteRoot
. (Join-Path $PSScriptRoot 'ShowcaseEvidence.Helpers.ps1')
$catalog = Get-Content -LiteralPath (Join-Path $websiteRoot 'data/showcase.json') -Raw | ConvertFrom-Json -Depth 40
$cards = @($catalog.cards | Where-Object walkthrough_url)
$sources = [ordered]@{}
$expected = [ordered]@{}
$pageRoot = Join-Path $websiteRoot 'content/showcase-examples'

foreach ($card in $cards) {
    if ($card.id -cnotmatch '^[a-z0-9]+(?:-[a-z0-9]+)*$' -or $sources.Contains($card.id)) {
        throw "Invalid or duplicate showcase walkthrough id: $($card.id)"
    }
    if ($card.walkthrough_url -cne "/showcase/$($card.id)/") { throw "Unexpected walkthrough route: $($card.walkthrough_url)" }
    $sourcePath = Resolve-ShowcasePath $repoRoot $card.source_path
    $sourceText = [IO.File]::ReadAllText($sourcePath).Replace("`r`n", "`n")
    $downloadPath = Resolve-ShowcasePath (Join-Path $websiteRoot 'static') $card.source_url.TrimStart('/')
    if (-not (Test-Path -LiteralPath $downloadPath -PathType Leaf) -or
        [IO.File]::ReadAllText($downloadPath).Replace("`r`n", "`n") -cne $sourceText) {
        throw "Refresh the generating source download with Build-ShowcaseEvidence.ps1 -ExampleId $($card.id)."
    }
    foreach ($related in $card.walkthrough.related) {
        if ($related -cnotin @($catalog.cards.id)) { throw "Unknown related example '$related' on '$($card.id)'." }
    }
    $sources[$card.id] = $sourceText
    $title = $card.title | ConvertTo-Json -Compress
    $description = $card.description | ConvertTo-Json -Compress
    $seoTitle = "$($card.title) in C# | OfficeIMO" | ConvertTo-Json -Compress
    $expected[(Join-Path $pageRoot "$($card.id).md")] = @"
---
# Generated from data/showcase.json by scripts/Sync-ShowcaseWalkthroughs.ps1.
title: $title
description: $description
layout: showcase-example
meta.showcase_id: $($card.id)
meta.seo_title: $seoTitle
meta.data_shortcode: showcase-walkthrough
meta.data_path: showcase
meta.data_mode: override
---

Browse [the showcase](/showcase/) for generated documents and source code.
"@ + "`n"
}
$expected[(Join-Path $websiteRoot 'data/showcase_sources.json')] = ($sources | ConvertTo-Json -Depth 4).Replace("`r`n", "`n") + "`n"
foreach ($file in Get-ChildItem -LiteralPath $pageRoot -Filter '*.md' -File -ErrorAction SilentlyContinue) {
    if (-not $expected.Contains($file.FullName)) { throw "Unmapped generated showcase page needs review: $($file.FullName)" }
}
foreach ($entry in $expected.GetEnumerator()) {
    $content = $entry.Value.Replace("`r`n", "`n")
    $current = if (Test-Path -LiteralPath $entry.Key -PathType Leaf) { [IO.File]::ReadAllText($entry.Key).Replace("`r`n", "`n") } else { $null }
    if ($current -ceq $content) { continue }
    if ($Check) { throw "Generated showcase content is stale: $($entry.Key). Run Sync-ShowcaseWalkthroughs.ps1." }
    New-Item -ItemType Directory -Path (Split-Path -Parent $entry.Key) -Force | Out-Null
    [IO.File]::WriteAllText($entry.Key, $content, [Text.UTF8Encoding]::new($false))
}
Write-Host "Showcase walkthroughs verified: $($cards.Count) pages with code from their generating C# files."
