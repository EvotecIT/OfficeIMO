function Sync-DocumentationPages {
    param(
        [Parameter(Mandatory)][string] $Source,
        [Parameter(Mandatory)][string] $Destination
    )

    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    Get-ChildItem -LiteralPath $Destination -Force -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -Recurse }

    foreach ($sourceFile in Get-ChildItem -LiteralPath $Source -File -Filter '*.md' | Sort-Object Name) {
        $slug = [System.IO.Path]::GetFileNameWithoutExtension($sourceFile.Name)
        if ($slug -eq '_index') {
            $targetPath = Join-Path $Destination 'index.md'
        } else {
            $targetDirectory = Join-Path $Destination $slug
            New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
            $targetPath = Join-Path $targetDirectory 'index.md'
        }

        $content = [IO.File]::ReadAllText($sourceFile.FullName)
        [IO.File]::WriteAllText($targetPath, $content, [Text.UTF8Encoding]::new($false))
    }
}

function Get-DocumentationTitle {
    param(
        [Parameter(Mandatory)][string] $Path,
        [Parameter(Mandatory)][string] $Fallback
    )

    $lines = @(Get-Content -LiteralPath $Path)
    if ($lines.Count -gt 0 -and $lines[0].Trim() -eq '---') {
        for ($index = 1; $index -lt $lines.Count; $index++) {
            $line = $lines[$index].Trim()
            if ($line -eq '---') {
                break
            }
            if ($line.StartsWith('title:', [StringComparison]::OrdinalIgnoreCase)) {
                $title = $line.Substring('title:'.Length).Trim().Trim('"').Trim("'")
                if (-not [string]::IsNullOrWhiteSpace($title)) {
                    return $title
                }
            }
        }
    }

    return $Fallback
}

function Sync-DocumentationToc {
    param(
        [Parameter(Mandatory)][string] $Source,
        [Parameter(Mandatory)][string] $TocPath
    )

    $toc = @(Get-Content -LiteralPath $TocPath -Raw | ConvertFrom-Json)
    $sections = @($toc | Where-Object title -EQ 'PSWriteOffice')
    if ($sections.Count -ne 1) {
        throw "The OfficeIMO documentation TOC must contain exactly one PSWriteOffice section."
    }

    $sourcePages = foreach ($sourceFile in Get-ChildItem -LiteralPath $Source -File -Filter '*.md' | Sort-Object Name) {
        $sourceSlug = [System.IO.Path]::GetFileNameWithoutExtension($sourceFile.Name)
        $key = if ($sourceSlug -eq '_index') { '__index__' } else { $sourceSlug }
        $fallbackTitle = if ($sourceSlug -eq '_index') {
            'PSWriteOffice Overview'
        } else {
            [System.Globalization.CultureInfo]::InvariantCulture.TextInfo.ToTitleCase($sourceSlug.Replace('-', ' '))
        }
        $targetPath = if ($sourceSlug -eq '_index') {
            'pswriteoffice/index.md'
        } else {
            "pswriteoffice/$sourceSlug/index.md"
        }
        $href = if ($sourceSlug -eq '_index') {
            '/docs/pswriteoffice/'
        } else {
            "/docs/pswriteoffice/$sourceSlug/"
        }

        [PSCustomObject]@{
            Key = $key
            Title = Get-DocumentationTitle -Path $sourceFile.FullName -Fallback $fallbackTitle
            Path = $targetPath
            Href = $href
        }
    }

    $pageByKey = @{}
    foreach ($page in $sourcePages) {
        $pageByKey[$page.Key] = $page
    }

    $usedKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($existing in @($sections[0].items)) {
        $existingPath = ([string] $existing.path).Replace('\', '/')
        $key = if ($existingPath -eq 'pswriteoffice/index.md') {
            '__index__'
        } elseif ($existingPath -match '^pswriteoffice/(?<slug>[^/]+)/index\.md$') {
            $Matches['slug']
        } else {
            $null
        }

        if ($key -and $pageByKey.ContainsKey($key) -and $usedKeys.Add($key)) {
            $existing.title = $pageByKey[$key].Title
            $items.Add($existing)
        }
    }

    foreach ($page in @($sourcePages | Where-Object { -not $usedKeys.Contains($_.Key) } | Sort-Object Title, Key)) {
        $items.Add([PSCustomObject]@{
            title = $page.Title
            path = $page.Path
            href = $page.Href
        })
        [void] $usedKeys.Add($page.Key)
    }

    $sections[0].items = $items.ToArray()
    $toc | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $TocPath -Encoding utf8
}

function Set-PSWriteOfficeDocumentationSnapshot {
    param(
        [Parameter(Mandatory)][string] $SiteRootPath,
        [Parameter(Mandatory)] $Source
    )

    $repository = [regex]::Escape($Source.Repo)
    $module = Import-PowerShellDataFile -LiteralPath (Join-Path $SiteRootPath 'data/apidocs/powershell/PSWriteOffice.psd1')
    $sourceRef = [Uri]::EscapeDataString($Source.Ref)
    # Only floating repository refs change. Explicit historical tags, other
    # repositories, URL suffixes and branches such as main-next retain their meaning.
    $pattern = '(?<prefix>https://(?:github\.com/' + $repository +
        '/(?:blob|tree)/|raw\.githubusercontent\.com/' + $repository +
        '/))(?:main|master)(?=[/?#\s"''<>)]|$)'
    $replacement = [System.Text.RegularExpressions.MatchEvaluator] {
        param($match)
        $match.Groups['prefix'].Value + $sourceRef
    }
    $catalogPath = Join-Path $SiteRootPath 'data/pswriteoffice_command_catalog.json'
    $catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json -Depth 20
    # A family's examples belong to this module version even if its catalog was
    # previously pinned. Historical links in guide prose are handled separately.
    $examplesPattern = '^(?<prefix>https://github\.com/' + $repository +
        '/tree/)[^/]+(?=/PowerShell/PSWriteOffice/Examples(?:[/#?]|$))'
    $catalogChanged = $false
    foreach ($family in @($catalog.families)) {
        $examplesUrl = [regex]::Replace([string] $family.examplesUrl, $examplesPattern, $replacement)
        if ($examplesUrl -cne $family.examplesUrl) {
            $family.examplesUrl = $examplesUrl
            $catalogChanged = $true
        }
    }
    if ($catalogChanged) {
        [IO.File]::WriteAllText($catalogPath, ($catalog | ConvertTo-Json -Depth 20 -Compress) + [Environment]::NewLine)
    }

    $pagesChanged = $false
    $documentationRoot = Join-Path $SiteRootPath 'content/docs/pswriteoffice'
    foreach ($page in Get-ChildItem -LiteralPath $documentationRoot -Filter '*.md' -File -Recurse) {
        $content = [IO.File]::ReadAllText($page.FullName)
        $pinnedContent = [regex]::Replace($content, $pattern, $replacement)
        $family = @($catalog.families | Where-Object id -CEQ $page.Directory.Name)
        if ($page.Name -ceq 'index.md' -and $family.Count -eq 1) {
            # Imported and fallback summaries follow the same catalog totals.
            $summaryCount = [regex]::new('\b\d+(?= exported commands\b)|(?<=\bexports )\d+(?= commands\b)')
            $pinnedContent = $summaryCount.Replace($pinnedContent, [string] $family[0].commandCount, 1)
        }
        if ($content -cne $pinnedContent) {
            [IO.File]::WriteAllText($page.FullName, $pinnedContent, [Text.UTF8Encoding]::new($false))
            $pagesChanged = $true
        }
    }
    [PSCustomObject]@{ CatalogUpdated = $catalogChanged; DocumentationUpdated = $pagesChanged }
}

function New-DocumentationPageEvidence {
    param([Parameter(Mandatory)][string] $RepositoryRoot)

    $sourceRoot = Join-Path $RepositoryRoot 'Website/content/project-docs/docs'
    foreach ($page in Get-ChildItem -LiteralPath $sourceRoot -File -Filter '*.md' | Sort-Object Name) {
        $sitePath = if ($page.BaseName -ceq '_index') { 'content/docs/pswriteoffice/index.md' }
            else { "content/docs/pswriteoffice/$($page.BaseName)/index.md" }
        [ordered]@{
            sourcePath = 'Website/content/project-docs/docs/' + $page.Name
            source = Get-DocumentationCatalogFingerprint -Path $page.FullName
            sitePath = $sitePath
            derived = $null
        }
    }
}

function Get-PSWriteOfficeNavigationFingerprint {
    param([Parameter(Mandatory)][string] $SiteRootPath)

    $toc = Get-Content -LiteralPath (Join-Path $SiteRootPath 'content/docs/toc.json') -Raw | ConvertFrom-Json
    $sections = @($toc | Where-Object title -CEQ 'PSWriteOffice')
    if ($sections.Count -ne 1) { throw 'Exactly one PSWriteOffice navigation section is required.' }
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($sections[0] | ConvertTo-Json -Depth 8 -Compress))
    $hash = [Security.Cryptography.SHA256]::Create()
    try { [ordered]@{ bytes = $bytes.Length; sha256 = [Convert]::ToHexString($hash.ComputeHash($bytes)).ToLowerInvariant() } }
    finally { $hash.Dispose() }
}

function Assert-DocumentationPageSnapshot {
    param(
        [Parameter(Mandatory)][string] $SiteRootPath,
        [Parameter(Mandatory)] $Evidence
    )

    $pages = @($Evidence.documentation)
    if ($pages.Count -eq 0) { throw 'The documentation snapshot has no page evidence.' }
    $expectedPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($page in $pages) {
        # The owning producer imports flat Markdown source files into page folders.
        $fileName = [IO.Path]::GetFileName([string] $page.sourcePath)
        if ($page.sourcePath -cne ('Website/content/project-docs/docs/' + $fileName) -or
            $fileName -notmatch '^[a-zA-Z0-9_-]+\.md$' -or $page.source.bytes -le 0 -or
            $page.source.sha256 -cnotmatch '^[0-9a-f]{64}$') { throw 'Invalid documentation source evidence.' }
        $slug = [IO.Path]::GetFileNameWithoutExtension($fileName)
        $sitePath = if ($slug -ceq '_index') { 'content/docs/pswriteoffice/index.md' }
            else { "content/docs/pswriteoffice/$slug/index.md" }
        if ($page.sitePath -cne $sitePath -or -not $expectedPaths.Add($sitePath)) { throw 'Invalid documentation page mapping.' }
        $actual = Get-DocumentationCatalogFingerprint -Path (Join-Path $SiteRootPath $sitePath)
        if ($actual.bytes -ne $page.derived.bytes -or $actual.sha256 -cne $page.derived.sha256) { throw "Documentation page differs from its evidence: '$sitePath'." }
    }
    $actualPages = @(Get-ChildItem -LiteralPath (Join-Path $SiteRootPath 'content/docs/pswriteoffice') -File -Recurse -Filter '*.md')
    if ($actualPages.Count -ne $expectedPaths.Count) { throw 'The documentation page inventory differs from its evidence.' }
    foreach ($page in $actualPages) {
        $relative = [IO.Path]::GetRelativePath($SiteRootPath, $page.FullName).Replace('\', '/')
        if (-not $expectedPaths.Contains($relative)) { throw 'The documentation page inventory differs from its evidence.' }
    }
    $navigation = Get-PSWriteOfficeNavigationFingerprint -SiteRootPath $SiteRootPath
    if ($navigation.bytes -ne $Evidence.navigation.bytes -or $navigation.sha256 -cne $Evidence.navigation.sha256) {
        throw 'The PSWriteOffice navigation differs from its evidence.'
    }
}
