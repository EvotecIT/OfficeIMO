[CmdletBinding()]
param([switch] $Check)

<#
.SYNOPSIS
Derives the links between the API reference and the guides and examples that use it.

.DESCRIPTION
Reads the merged xref map (data/xrefmap.json) for every documented type, member, and
PowerShell command, then scans docs and workflow pages (inline code and C#/PowerShell
blocks) and the showcase example sources for the symbols they use. It writes:

- data/apidocs/related/<package>.json: "Guides & Samples" manifests for each API package,
  attached to the apidocs pipeline steps through relatedContentManifests.
- data/api_links.json: the API symbols each page uses, so the page can link its code
  to the reference and list the API it relies on.

Nothing here is authored by hand; rerun after content or API changes. -Check fails when
the committed output is stale.
#>

$ErrorActionPreference = 'Stop'
$websiteRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent $websiteRoot

# Pages that teach the API. Comparisons, blog posts, and browser tool pages are left out:
# they mention types in passing or describe other products.
$collections = @(
    @{ Input = 'content/docs'; Output = '/docs' }
    @{ Input = 'content/pdf-workflows'; Output = '/pdf' }
    @{ Input = 'content/conversions'; Output = '/convert' }
    @{ Input = 'content/solutions'; Output = '/solutions' }
    @{ Input = 'content/products'; Output = '/products' }
)
$codeLanguages = @('cs', 'csharp', 'c#', 'powershell', 'ps1', 'pwsh', 'ps')
$maxPagesPerKind = [ordered]@{ guide = 6; sample = 4 }
$maxLinksPerPage = 40
$summaryLimit = 180

function Get-FrontMatter([string] $text) {
    $meta = @{}
    $body = $text
    if ($text -match '(?s)\A---\n(.*?)\n---\n?(.*)\z') {
        $body = $Matches[2]
        foreach ($line in $Matches[1] -split "`n") {
            if ($line -match '^([A-Za-z_][\w.]*):\s*(.*)$') {
                $key = $Matches[1]
                $value = $Matches[2].Trim()
                if ($value -match '^"(.*)"$') { $value = $Matches[1] -replace '\\"', '"' }
                elseif ($value -match "^'(.*)'$") { $value = $Matches[1] -replace "''", "'" }
                $meta[$key] = $value
            }
        }
    }
    [pscustomobject]@{ Meta = $meta; Body = $body }
}

function Get-Route([string] $file, [string] $inputRoot, [string] $outputRoot, [hashtable] $meta) {
    $relative = [IO.Path]::GetRelativePath($inputRoot, $file).Replace('\', '/')
    $segments = [Collections.Generic.List[string]]::new()
    foreach ($segment in ($relative -replace '\.md$', '') -split '/') { $segments.Add($segment) }
    if ($segments[$segments.Count - 1] -eq 'index' -or $segments[$segments.Count - 1] -eq '_index') { $segments.RemoveAt($segments.Count - 1) }
    if ($meta.slug -and $segments.Count -gt 0) { $segments[$segments.Count - 1] = $meta.slug }
    $path = ($outputRoot.TrimEnd('/') + '/' + ($segments -join '/')).TrimEnd('/') + '/'
    $path -replace '/+', '/'
}

function Get-Summary([string] $value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    $clean = ($value -replace '\s+', ' ').Trim()
    if ($clean.Length -le $summaryLimit) { return $clean }
    $cut = $clean.Substring(0, $summaryLimit)
    $space = $cut.LastIndexOf(' ')
    if ($space -gt 100) { $cut = $cut.Substring(0, $space) }
    $cut.TrimEnd(',', ';', ':', '.') + '...'
}

# Symbol index from the merged xref map.
$xref = [Text.Json.JsonDocument]::Parse([IO.File]::ReadAllText((Join-Path $websiteRoot 'data/xrefmap.json')))
$typesByFullName = @{}
$typesByShortName = @{}
$members = @{}
$commands = @{}
foreach ($reference in $xref.RootElement.GetProperty('references').EnumerateArray()) {
    $uid = $reference.GetProperty('uid').GetString()
    $href = $reference.GetProperty('href').GetString()
    $name = $reference.GetProperty('name').GetString()
    if (-not $href.StartsWith('/api/')) { continue }
    $package = $href.Split('/')[2]
    if ($uid -match '^[MPFE]:(.+?)(\(.*\))?$') {
        $key = $Matches[1] -replace '``?\d+', ''
        if ($key.EndsWith('.#ctor') -or $members.ContainsKey($key)) { continue }
        $members[$key] = [pscustomobject]@{ Name = $name; Href = $href }
        continue
    }
    if ($href.Contains('#')) { continue }
    if ($package -eq 'powershell') {
        if ($uid -cmatch '^[A-Z][a-z]+-[A-Za-z0-9]+$') {
            $commands[$uid] = [pscustomobject]@{ Uid = $uid; Name = $uid; Href = $href; Package = $package; Namespace = '' }
        }
        continue
    }
    $short = ($name -replace '<.*$', '') -replace '``?\d+$', ''
    $namespace = if ($uid.Length -gt $short.Length -and $uid.LastIndexOf('.') -gt 0) { $uid.Substring(0, $uid.LastIndexOf('.')) } else { '' }
    $type = [pscustomobject]@{ Uid = $uid; Name = $short; Href = $href; Package = $package; Namespace = $namespace }
    $typesByFullName[($uid -replace '``?\d+$', '')] = $type
    if (-not $typesByShortName.ContainsKey($short)) { $typesByShortName[$short] = [Collections.Generic.List[object]]::new() }
    $typesByShortName[$short].Add($type)
}
$xref.Dispose()

# API steps from the pipeline: package base URLs, docs homes, and curated manifests.
$pipelinePath = Join-Path $websiteRoot 'pipeline.json'
$pipeline = Get-Content -LiteralPath $pipelinePath -Raw | ConvertFrom-Json -Depth 40
# Multi-input suite steps (the adapters) have no single package page to attach guides to.
$apiSteps = @($pipeline.steps | Where-Object { $_.task -eq 'apidocs' -and $_.baseUrl })
$packageByDocsHome = @{}
foreach ($step in $apiSteps) {
    if ($step.docsHome -and $step.baseUrl) { $packageByDocsHome[$step.docsHome] = $step.baseUrl.Split('/')[2] }
}

# A short name counts only when the page imports its namespace, belongs to its package, or
# the name is a distinctive compound: a Word guide's `Paragraph` is not the Markdown type.
function Resolve-Type([string] $short, [string[]] $usings, [string] $contextPackage) {
    $candidates = $typesByShortName[$short]
    if (-not $candidates) { return $null }
    $byUsing = @($candidates | Where-Object { $usings -contains $_.Namespace })
    if ($byUsing.Count -eq 1) { return $byUsing[0] }
    $byPackage = @($candidates | Where-Object Package -eq $contextPackage)
    if ($byPackage.Count -eq 1) { return $byPackage[0] }
    $distinctive = $short -cmatch '^[A-Z][a-z0-9]+(?:[A-Z][A-Za-z0-9]*)+$' -and $short.Length -ge 8
    if ($candidates.Count -eq 1 -and $distinctive) { return $candidates[0] }
    $null
}

# String literals and comments hold file names, cell addresses, and prose, not API use.
function Remove-CodeNoise([string] $code, [string] $language) {
    $pattern = if ($language -in @('powershell', 'ps1', 'pwsh', 'ps')) {
        "'(?:[^']|'')*'|`"(?:``.|[^`"``])*`"|<#.*?#>|#[^\n]*"
    } else {
        '@"(?:[^"]|"")*"|\$?"(?:\\.|[^"\\\n])*"|''(?:\\.|[^''\\\n])''|/\*.*?\*/|//[^\n]*'
    }
    [regex]::Replace($code, $pattern, ' ', [Text.RegularExpressions.RegexOptions]::Singleline)
}

function Add-Mention([hashtable] $mentions, $symbol, [string] $kind, [int] $weight) {
    $key = "$kind|$($symbol.Href)"
    if (-not $mentions.ContainsKey($key)) {
        $mentions[$key] = [pscustomobject]@{ Kind = $kind; Symbol = $symbol; Weight = 0 }
    }
    $mentions[$key].Weight += $weight
}

function Find-Mentions([string] $code, [string[]] $usings, [string] $contextPackage, [hashtable] $mentions, [int] $weight) {
    foreach ($match in [regex]::Matches($code, '\b[A-Z][a-z]+-[A-Z][A-Za-z0-9]+\b')) {
        if ($commands.ContainsKey($match.Value)) { Add-Mention $mentions $commands[$match.Value] 'command' $weight }
    }
    foreach ($match in [regex]::Matches($code, '(?<![\w.\-$@])[A-Z][A-Za-z0-9_]*(?:\.[A-Z][A-Za-z0-9_]*)*')) {
        $parts = $match.Value.Split('.')
        $type = $null
        $memberIndex = -1
        for ($i = $parts.Count; $i -ge 1 -and -not $type; $i--) {
            $candidate = ($parts[0..($i - 1)] -join '.')
            if ($typesByFullName.ContainsKey($candidate)) { $type = $typesByFullName[$candidate]; $memberIndex = $i }
        }
        if (-not $type) {
            for ($i = 0; $i -lt $parts.Count -and -not $type; $i++) {
                $type = Resolve-Type $parts[$i] $usings $contextPackage
                $memberIndex = $i + 1
            }
        }
        if (-not $type) { continue }
        Add-Mention $mentions $type 'type' $weight
        if ($memberIndex -lt $parts.Count) {
            $fullType = $type.Uid -replace '``?\d+$', ''
            $member = $members["$fullType.$($parts[$memberIndex])"]
            if ($member) { Add-Mention $mentions $member 'member' $weight }
        }
    }
}

function Get-PageMentions([string] $body, [string] $contextPackage) {
    $mentions = @{}
    $normalized = $body.Replace("`r`n", "`n")
    $blocks = [regex]::Matches($normalized, '(?ms)^[ \t]*```[ \t]*([\w#+-]*)[^\n]*\n(.*?)^[ \t]*```')
    $usings = @([regex]::Matches($normalized, '(?m)^\s*using\s+(?:static\s+)?([A-Za-z_][\w.]*)\s*;') | ForEach-Object { $_.Groups[1].Value })
    foreach ($block in $blocks) {
        if ($codeLanguages -contains $block.Groups[1].Value.ToLowerInvariant()) {
            $language = $block.Groups[1].Value.ToLowerInvariant()
            Find-Mentions (Remove-CodeNoise $block.Groups[2].Value $language) $usings $contextPackage $mentions 1
        }
    }
    # Prose names an API on purpose; weigh inline code above incidental use in a sample.
    $prose = [regex]::Replace($normalized, '(?ms)^[ \t]*```.*?^[ \t]*```', '')
    foreach ($inline in [regex]::Matches($prose, '`([^`\n]+)`')) {
        # Very short spans such as `A1` are cell addresses or values, not type names.
        if ($inline.Groups[1].Value.Trim().Length -lt 4) { continue }
        Find-Mentions $inline.Groups[1].Value $usings $contextPackage $mentions 3
    }
    $mentions
}

$pages = [Collections.Generic.List[object]]::new()
foreach ($collection in $collections) {
    $inputRoot = Join-Path $websiteRoot $collection.Input
    foreach ($file in Get-ChildItem -LiteralPath $inputRoot -Recurse -Filter '*.md' | Sort-Object FullName) {
        $parsed = Get-FrontMatter ([IO.File]::ReadAllText($file.FullName).Replace("`r`n", "`n"))
        if ($parsed.Meta.draft -eq 'true' -or -not $parsed.Meta.title) { continue }
        $route = Get-Route $file.FullName $inputRoot $collection.Output $parsed.Meta
        $contextPackage = $null
        foreach ($docsHome in $packageByDocsHome.Keys) { if ($route.StartsWith($docsHome)) { $contextPackage = $packageByDocsHome[$docsHome] } }
        $mentions = Get-PageMentions $parsed.Body $contextPackage
        if ($mentions.Count -eq 0) { continue }
        $pages.Add([pscustomobject]@{ Route = $route; Title = $parsed.Meta.title; Summary = Get-Summary $parsed.Meta.description; Kind = 'guide'; Mentions = $mentions })
    }
}

$showcase = Get-Content -LiteralPath (Join-Path $websiteRoot 'data/showcase.json') -Raw | ConvertFrom-Json -Depth 40
foreach ($card in @($showcase.cards | Where-Object { $_.walkthrough_url -and $_.source_path } | Sort-Object id)) {
    $sourcePath = Join-Path $repoRoot $card.source_path
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "Showcase source not found: $($card.source_path)" }
    $source = [IO.File]::ReadAllText($sourcePath).Replace("`r`n", "`n")
    $fence = '```'
    $mentions = Get-PageMentions "$fence`csharp`n$source`n$fence`n" $null
    if ($mentions.Count -eq 0) { continue }
    $pages.Add([pscustomobject]@{ Route = $card.walkthrough_url; Title = $card.title; Summary = Get-Summary $card.description; Kind = 'sample'; Mentions = $mentions })
}

# Each page's links, strongest first, for the page to link its code and list its API.
$pageLinks = [ordered]@{}
foreach ($page in $pages | Sort-Object Route) {
    $links = @($page.Mentions.Values |
        Sort-Object @{ Expression = 'Weight'; Descending = $true }, @{ Expression = { $_.Symbol.Name } }, @{ Expression = { $_.Symbol.Href } } |
        Select-Object -First $maxLinksPerPage |
        ForEach-Object {
            # Templates write these into a JSON script element and href attributes unescaped.
            if ($_.Symbol.Name -notmatch '^[\w.-]+$' -or $_.Symbol.Href -notmatch '^/api/[\w./#-]+$') { throw "Unsafe API link: $($_.Symbol.Name) $($_.Symbol.Href)" }
            [ordered]@{ n = $_.Symbol.Name; h = $_.Symbol.Href; k = $_.Kind }
        })
    $pageLinks[$page.Route] = $links
}

# Guides & Samples per API package: each type keeps the pages that use it most.
$pagesByType = @{}
foreach ($page in $pages) {
    foreach ($mention in $page.Mentions.Values) {
        if ($mention.Kind -eq 'member') { continue }
        $uid = $mention.Symbol.Uid
        if (-not $pagesByType.ContainsKey($uid)) { $pagesByType[$uid] = [Collections.Generic.List[object]]::new() }
        $pagesByType[$uid].Add([pscustomobject]@{ Page = $page; Weight = $mention.Weight; Symbol = $mention.Symbol })
    }
}
$targetsByPackage = @{}
foreach ($uid in $pagesByType.Keys) {
    # Guides and examples get separate slots so a type's reference shows both.
    $ranked = foreach ($kind in $maxPagesPerKind.Keys) {
        $pagesByType[$uid] | Where-Object { $_.Page.Kind -eq $kind } |
            Sort-Object @{ Expression = 'Weight'; Descending = $true }, @{ Expression = { $_.Page.Route } } |
            Select-Object -First $maxPagesPerKind[$kind]
    }
    foreach ($item in $ranked) {
        $package = $item.Symbol.Package
        if (-not $targetsByPackage.ContainsKey($package)) { $targetsByPackage[$package] = @{} }
        $byRoute = $targetsByPackage[$package]
        if (-not $byRoute.ContainsKey($item.Page.Route)) { $byRoute[$item.Page.Route] = [pscustomobject]@{ Page = $item.Page; Targets = [Collections.Generic.List[string]]::new(); Weight = 0 } }
        $byRoute[$item.Page.Route].Targets.Add($uid)
        $byRoute[$item.Page.Route].Weight += $item.Weight
    }
}

$expected = [ordered]@{}
$manifestRoot = Join-Path $websiteRoot 'data/apidocs/related'
$stepManifests = @{}
foreach ($step in $apiSteps) {
    $package = $step.baseUrl.Split('/')[2]
    if (-not $targetsByPackage.ContainsKey($package)) { continue }
    $curated = @{}
    foreach ($path in @($step.relatedContentManifests) + @($step.relatedContentManifest) | Where-Object { $_ -and $_ -notlike './data/apidocs/related/*' }) {
        foreach ($entry in @(Get-Content -LiteralPath (Join-Path $websiteRoot $path) -Raw | ConvertFrom-Json)) {
            foreach ($target in $entry.targets) { $curated["$($entry.url)|$target"] = $true }
        }
    }
    $entries = foreach ($item in $targetsByPackage[$package].Values |
        Sort-Object @{ Expression = { $_.Page.Kind -ne 'guide' } }, @{ Expression = 'Weight'; Descending = $true }, @{ Expression = { $_.Page.Route } }) {
        $targets = @($item.Targets | Where-Object { -not $curated.ContainsKey("$($item.Page.Route)|$_") } | Sort-Object -Unique)
        if ($targets.Count -eq 0) { continue }
        $entry = [ordered]@{ title = $item.Page.Title; url = $item.Page.Route }
        if ($item.Page.Summary) { $entry.summary = $item.Page.Summary }
        $entry.kind = $item.Page.Kind
        $entry.targets = $targets
        '  ' + ($entry | ConvertTo-Json -Compress -Depth 4)
    }
    if (-not $entries) { continue }
    $relative = "./data/apidocs/related/$package.json"
    $stepManifests[$step.id] = $relative
    $expected[(Join-Path $manifestRoot "$package.json")] = "[`n" + (@($entries) -join ",`n") + "`n]`n"
}

$linkLines = foreach ($route in $pageLinks.Keys) {
    '    ' + ($route | ConvertTo-Json -Compress) + ': ' + (ConvertTo-Json -InputObject @($pageLinks[$route]) -Compress -Depth 4)
}
$expected[(Join-Path $websiteRoot 'data/api_links.json')] = "{`n  `"pages`": {`n" + (@($linkLines) -join ",`n") + "`n  }`n}`n"

# The apidocs steps read the curated manifests first, then the derived one.
$pipelineErrors = foreach ($step in $apiSteps) {
    $configured = @($step.relatedContentManifests) + @($step.relatedContentManifest) | Where-Object { $_ }
    $derived = $stepManifests[$step.id]
    $configuredDerived = @($configured | Where-Object { $_ -like './data/apidocs/related/*' })
    if ($derived -and $configuredDerived -notcontains $derived) { "$($step.id) does not read $derived" }
    foreach ($path in $configuredDerived) { if ($path -ne $derived) { "$($step.id) reads $path, which is no longer generated" } }
}

$stale = @()
foreach ($path in $expected.Keys) {
    $current = if (Test-Path -LiteralPath $path) { [IO.File]::ReadAllText($path).Replace("`r`n", "`n") } else { $null }
    if ($current -cne $expected[$path]) { $stale += $path }
}
$orphans = @(Get-ChildItem -LiteralPath $manifestRoot -Filter '*.json' -ErrorAction SilentlyContinue | Where-Object { -not $expected.Contains($_.FullName) })

if ($Check) {
    $problems = @($stale | ForEach-Object { "stale: $([IO.Path]::GetRelativePath($websiteRoot, $_))" }) + @($orphans | ForEach-Object { "orphan: $($_.Name)" }) + @($pipelineErrors)
    if ($problems.Count -gt 0) { throw "API cross-links are out of date; run scripts/Build-ApiCrossLinks.ps1.`n" + ($problems -join "`n") }
    Write-Host "API cross-links are current: $($pages.Count) pages, $($stepManifests.Count) package manifests."
    return
}

New-Item -ItemType Directory -Force -Path $manifestRoot | Out-Null
foreach ($path in $stale) { [IO.File]::WriteAllText($path, $expected[$path], [Text.UTF8Encoding]::new($false)) }
foreach ($orphan in $orphans) { Remove-Item -LiteralPath $orphan.FullName }
if ($pipelineErrors) { Write-Warning ("Update pipeline.json relatedContentManifests:`n" + ($pipelineErrors -join "`n")) }
Write-Host "API cross-links: $($pages.Count) pages, $($stepManifests.Count) package manifests, $($stale.Count) files updated."
