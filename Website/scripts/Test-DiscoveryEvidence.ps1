param(
    [string] $SiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path,
    [string] $ArtifactRoot
)

$ErrorActionPreference = 'Stop'
$failures = [System.Collections.Generic.List[string]]::new()

function Add-Failure {
    param([string] $Message)
    $failures.Add($Message)
}

$siteConfig = Get-Content -LiteralPath (Join-Path $SiteRoot 'site.json') -Raw | ConvertFrom-Json
$signals = $siteConfig.AgentReadiness.ContentSignals
if (-not $signals.Enabled -or -not $signals.Search -or -not $signals.AiInput -or -not $signals.AiTrain) {
    Add-Failure 'AgentReadiness content signals must explicitly allow search, AI input, and AI training.'
}

$requiredBots = @('GPTBot', 'ChatGPT-User', 'OAI-SearchBot')
foreach ($bot in $requiredBots) {
    $rule = @($siteConfig.AgentReadiness.BotRules | Where-Object UserAgent -EQ $bot)
    if ($rule.Count -ne 1 -or [string] $rule[0].Allow -ne '/') {
        Add-Failure "AgentReadiness must explicitly allow $bot at the site root."
    }
}

$expectedSitemapPolicies = @{
    pages       = 'sourceDate'
    products    = 'sourceDate'
    conversions = 'sourceDate'
    solutions   = 'sourceDate'
    comparisons = 'sourceDate'
    docs        = 'sourceDate'
    blog        = 'publishedDate'
}
foreach ($entry in $expectedSitemapPolicies.GetEnumerator()) {
    $collection = @($siteConfig.Collections | Where-Object Name -EQ $entry.Key)
    if ($collection.Count -ne 1 -or [string] $collection[0].SitemapLastModified -ne $entry.Value) {
        Add-Failure "Collection '$($entry.Key)' must use sitemap last-modified policy '$($entry.Value)'."
    }
}

$requiredProductRoutes = @('/comparison/', '/comparisons/**')
foreach ($route in $requiredProductRoutes) {
    $bundle = @($siteConfig.AssetRegistry.RouteBundles | Where-Object Match -EQ $route)
    if ($bundle.Count -ne 1 -or @($bundle[0].Bundles) -notcontains 'product') {
        Add-Failure "Route '$route' must load the product visual bundle."
    }
}

$registryPath = Join-Path $SiteRoot 'data\comparison_evidence.json'
$registry = Get-Content -LiteralPath $registryPath -Raw | ConvertFrom-Json
if ([int] $registry.schemaVersion -ne 1) {
    Add-Failure 'comparison_evidence.json must use schemaVersion 1.'
}

$reviewed = [datetime]::MinValue
if (-not [datetime]::TryParseExact(
        [string] $registry.lastReviewed,
        'yyyy-MM-dd',
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None,
        [ref] $reviewed)) {
    Add-Failure 'comparison_evidence.json lastReviewed must use yyyy-MM-dd.'
}

$hub = Get-Content -LiteralPath (Join-Path $SiteRoot 'content\pages\comparisons.md') -Raw
$hubLayout = [regex]::Match($hub, '(?m)^layout:\s*(?<value>[^\r\n]+)\s*$').Groups['value'].Value.Trim()
if ($hubLayout -ne 'comparison-hub') {
    Add-Failure "The comparison hub must use the 'comparison-hub' layout."
}

$hubCardCount = [regex]::Matches($hub, 'class="imo-comparison-card(?:\s|")').Count
$registeredComparisonCount = @($registry.comparisons).Count
if ($hubCardCount -ne $registeredComparisonCount) {
    Add-Failure "The comparison hub has $hubCardCount visual cards for $registeredComparisonCount registered comparisons."
}

$hubLayoutPath = Join-Path $SiteRoot 'themes\officeimo\layouts\comparison-hub.html'
if (-not (Test-Path -LiteralPath $hubLayoutPath)) {
    Add-Failure "The comparison hub layout is missing at '$hubLayoutPath'."
} else {
    $hubLayoutContent = Get-Content -LiteralPath $hubLayoutPath -Raw
    if ($hubLayoutContent -notmatch [regex]::Escape('{{ data.comparison_evidence.comparisons.size }}')) {
        Add-Failure 'The comparison decision map count must render from comparison_evidence.json.'
    }
    if ($hubLayoutContent -notmatch [regex]::Escape('{{ data.comparison_evidence.lastReviewed }}')) {
        Add-Failure 'The comparison decision map review date must render from comparison_evidence.json.'
    }
}

$seenIds = @{}
$seenRoutes = @{}
foreach ($comparison in @($registry.comparisons)) {
    $id = [string] $comparison.id
    $route = [string] $comparison.route
    if ([string]::IsNullOrWhiteSpace($id) -or $seenIds.ContainsKey($id)) {
        Add-Failure "Comparison id '$id' is empty or duplicated."
    } else {
        $seenIds[$id] = $true
    }
    if ([string]::IsNullOrWhiteSpace($route) -or $seenRoutes.ContainsKey($route)) {
        Add-Failure "Comparison route '$route' is empty or duplicated."
    } else {
        $seenRoutes[$route] = $true
    }

    $sourcePath = Join-Path $SiteRoot ([string] $comparison.sourcePath)
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        Add-Failure "Comparison '$id' points to missing source '$($comparison.sourcePath)'."
        continue
    }

    $content = Get-Content -LiteralPath $sourcePath -Raw
    $descriptionMatch = [regex]::Match(
        $content,
        '(?m)^description:\s*(?:"(?<value>[^"]+)"|''(?<value>[^'']+)''|(?<value>[^\r\n]+))\s*$'
    )
    $description = $descriptionMatch.Groups['value'].Value.Trim()
    if ($description.Length -lt 120 -or $description.Length -gt 160) {
        Add-Failure "Comparison '$id' description is $($description.Length) characters; expected 120-160."
    }
    if ($content -match '(?m)^#\s+') {
        Add-Failure "Comparison '$id' contains a body H1 even though the layout renders the page title."
    }
    if ($content -notmatch '(?im)^##\s+(?:Choose|Choosing|When|Where)') {
        Add-Failure "Comparison '$id' must include an explicit decision section beginning with Choose, Choosing, When, or Where."
    }

    $checked = [datetime]::MinValue
    if (-not [datetime]::TryParseExact(
            [string] $comparison.checkedAt,
            'yyyy-MM-dd',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None,
            [ref] $checked)) {
        Add-Failure "Comparison '$id' checkedAt must use yyyy-MM-dd."
    } elseif ($checked.Date -gt [datetime]::UtcNow.Date -or $checked.Date -lt [datetime]::UtcNow.Date.AddDays(-400)) {
        Add-Failure "Comparison '$id' has a future or stale checkedAt date '$($comparison.checkedAt)'."
    }

    $sources = @($comparison.sources)
    if ($sources.Count -lt 2) {
        Add-Failure "Comparison '$id' must cite at least two first-party sources."
    }
    foreach ($source in $sources) {
        $url = [string] $source.url
        if ($url -notmatch '^https://') {
            Add-Failure "Comparison '$id' has a non-HTTPS source '$url'."
        } elseif ($content -notmatch [regex]::Escape($url)) {
            Add-Failure "Comparison '$id' does not link its registered source '$url'."
        }
    }

    $evidenceRoutes = @($comparison.officeImoEvidence)
    if ($evidenceRoutes.Count -lt 2) {
        Add-Failure "Comparison '$id' must link at least two OfficeIMO evidence routes."
    }
    foreach ($evidenceRoute in $evidenceRoutes) {
        if ($content -notmatch [regex]::Escape([string] $evidenceRoute)) {
            Add-Failure "Comparison '$id' does not link its registered OfficeIMO evidence '$evidenceRoute'."
        }
    }

    if ($hub -notmatch [regex]::Escape($route)) {
        Add-Failure "The comparison hub does not link '$route'."
    }
}

if (-not [string]::IsNullOrWhiteSpace($ArtifactRoot)) {
    $resolvedArtifactRoot = (Resolve-Path -LiteralPath $ArtifactRoot).Path
    $discoveryRoutes = @($registry.comparisons | ForEach-Object { [string] $_.route }) +
        @(
            '/comparisons/',
            '/docs/pswriteoffice/compare-office-automation-options/',
            '/compatibility/',
            '/licensing/'
        )

    foreach ($fileName in @('llms.txt', 'llms-full.txt')) {
        $artifactPath = Join-Path $resolvedArtifactRoot $fileName
        if (-not (Test-Path -LiteralPath $artifactPath)) {
            Add-Failure "Generated discovery artifact '$fileName' is missing."
            continue
        }

        $artifactContent = Get-Content -LiteralPath $artifactPath -Raw
        foreach ($route in $discoveryRoutes) {
            if ($artifactContent -notmatch [regex]::Escape($route)) {
                Add-Failure "Generated '$fileName' does not include discovery route '$route'."
            }
        }
    }

    $hubArtifactPath = Join-Path $resolvedArtifactRoot 'comparisons\index.html'
    if (-not (Test-Path -LiteralPath $hubArtifactPath)) {
        Add-Failure 'The rendered comparison hub is missing.'
    } else {
        $hubArtifact = Get-Content -LiteralPath $hubArtifactPath -Raw
        if ($hubArtifact -notmatch "$registeredComparisonCount maintained comparisons") {
            Add-Failure "The rendered comparison hub does not show the registered count '$registeredComparisonCount'."
        }
        if ($hubArtifact -notmatch [regex]::Escape("Reviewed $($registry.lastReviewed)")) {
            Add-Failure "The rendered comparison hub does not show registry review date '$($registry.lastReviewed)'."
        }
    }

    $sitemapPath = Join-Path $resolvedArtifactRoot 'sitemap.xml'
    if (-not (Test-Path -LiteralPath $sitemapPath)) {
        Add-Failure 'The generated sitemap.xml is missing.'
    } else {
        [xml] $sitemap = Get-Content -LiteralPath $sitemapPath -Raw
        $freshnessRoutes = @($registry.comparisons | ForEach-Object { [string] $_.route }) +
            @('/comparisons/', '/licensing/')
        foreach ($route in $freshnessRoutes) {
            $canonical = "https://officeimo.com$route"
            $entry = @($sitemap.urlset.url | Where-Object { [string] $_.loc -eq $canonical })
            if ($entry.Count -ne 1) {
                Add-Failure "Generated sitemap does not contain exactly one '$canonical' entry."
            } elseif ([string]::IsNullOrWhiteSpace([string] $entry[0].lastmod)) {
                Add-Failure "Generated sitemap entry '$canonical' has no lastmod."
            }
        }
    }
}

if ($failures.Count -gt 0) {
    throw "Discovery evidence validation failed:`n - $($failures -join "`n - ")"
}

[PSCustomObject]@{
    ComparisonCount = @($registry.comparisons).Count
    TrainingAllowed = [bool] $signals.AiTrain
    SitemapPolicyCount = $expectedSitemapPolicies.Count
    Status = 'passed'
}
