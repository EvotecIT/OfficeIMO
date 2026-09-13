[CmdletBinding()]
param(
    [string] $SiteRoot = (Split-Path -Parent $PSScriptRoot),
    [string] $PublishedRoot
)

$ErrorActionPreference = 'Stop'

$siteRootPath = (Resolve-Path -LiteralPath $SiteRoot).Path
$contentRoot = Join-Path $siteRootPath 'content'
$layoutRoot = Join-Path $siteRootPath 'themes\officeimo\layouts'
$siteConfigPath = Join-Path $siteRootPath 'site.json'
$siteConfig = Get-Content -LiteralPath $siteConfigPath -Raw | ConvertFrom-Json
$displayKeys = @(
    'eyebrow',
    'outcome',
    'primary_label',
    'primary_url',
    'secondary_label',
    'secondary_url',
    'summary_title',
    'package',
    'package_url',
    'runtime',
    'limit',
    'related_label',
    'related_url'
)
$failures = [System.Collections.Generic.List[string]]::new()
$publishedRootPath = if ($PublishedRoot) {
    (Resolve-Path -LiteralPath $PublishedRoot).Path
}

foreach ($contentFile in Get-ChildItem -LiteralPath $contentRoot -Recurse -File -Filter '*.md') {
    $collection = $null
    $content = Get-Content -LiteralPath $contentFile.FullName -Raw
    $frontMatterMatch = [regex]::Match($content, '\A---\s*\r?\n(?<frontMatter>.*?)\r?\n---\s*\r?\n', 'Singleline')
    if (-not $frontMatterMatch.Success) {
        continue
    }

    $frontMatter = $frontMatterMatch.Groups['frontMatter'].Value
    $layout = [regex]::Match($frontMatter, '(?m)^layout:\s*(?<value>[^\r\n]+)\s*$').Groups['value'].Value.Trim()
    if (-not $layout) {
        $collection = $siteConfig.Collections | Where-Object {
            $collectionRoot = [IO.Path]::GetFullPath((Join-Path $siteRootPath $_.Input))
            $contentFile.FullName.StartsWith($collectionRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1
        $layout = [string] $collection.DefaultLayout
    }

    if (-not $collection) {
        $collection = $siteConfig.Collections | Where-Object {
            $candidateRoot = [IO.Path]::GetFullPath((Join-Path $siteRootPath $_.Input))
            $contentFile.FullName.StartsWith($candidateRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)
        } | Select-Object -First 1
    }

    if (-not $collection) {
        $failures.Add("$($contentFile.FullName): no content collection could be resolved.")
        continue
    }

    if (-not $layout) {
        $failures.Add("$($contentFile.FullName): no effective layout could be resolved.")
        continue
    }

    $layoutPath = Join-Path $layoutRoot ($layout + '.html')
    if (-not (Test-Path -LiteralPath $layoutPath)) {
        $failures.Add("$($contentFile.FullName): layout '$layout' does not exist at '$layoutPath'.")
        continue
    }

    $layoutContent = Get-Content -LiteralPath $layoutPath -Raw
    foreach ($key in $displayKeys) {
        $metadataMatch = [regex]::Match($frontMatter, "(?m)^meta\.$key\s*:\s*(?<value>[^\r\n]+)\s*$")
        if ($metadataMatch.Success -and
            $layoutContent -notmatch [regex]::Escape("page.meta.$key")) {
            $relativePath = [IO.Path]::GetRelativePath($siteRootPath, $contentFile.FullName)
            $failures.Add("$relativePath declares meta.$key, but layout '$layout' does not render it.")
        }

        if (-not $publishedRootPath -or -not $metadataMatch.Success) {
            continue
        }

        $collectionRoot = [IO.Path]::GetFullPath((Join-Path $siteRootPath $collection.Input))
        $relativeContentPath = [IO.Path]::GetRelativePath($collectionRoot, $contentFile.FullName)
        $relativeDirectory = Split-Path -Parent $relativeContentPath
        $slugMatch = [regex]::Match($frontMatter, '(?m)^slug:\s*(?<value>[^\r\n]+)\s*$')
        $slug = if ($slugMatch.Success) {
            $slugMatch.Groups['value'].Value.Trim().Trim('"', "'")
        } else {
            [IO.Path]::GetFileNameWithoutExtension($contentFile.Name).TrimStart('_')
        }
        $routeParts = @($collection.Output.Trim('/'))
        if ($relativeDirectory) {
            $routeParts += $relativeDirectory -split '[\\/]'
        }
        if ($slug -and $slug -ne 'index') {
            $routeParts += $slug
        }
        $publishedPath = Join-Path $publishedRootPath ((@($routeParts | Where-Object { $_ }) + 'index.html') -join [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $publishedPath)) {
            $relativePath = [IO.Path]::GetRelativePath($siteRootPath, $contentFile.FullName)
            $failures.Add("$relativePath declares meta.$key, but its published page was not found at '$publishedPath'.")
            continue
        }

        $metadataValue = $metadataMatch.Groups['value'].Value.Trim().Trim('"', "'")
        $publishedContent = Get-Content -LiteralPath $publishedPath -Raw
        $encodedValue = [Net.WebUtility]::HtmlEncode($metadataValue)
        if (-not $publishedContent.Contains($metadataValue, [StringComparison]::Ordinal) -and
            -not $publishedContent.Contains($encodedValue, [StringComparison]::Ordinal)) {
            $relativePath = [IO.Path]::GetRelativePath($siteRootPath, $contentFile.FullName)
            $failures.Add("$relativePath declares meta.$key, but its value is missing from '$publishedPath'.")
        }
    }
}

if ($failures.Count -gt 0) {
    throw "Content layout contract validation failed:`n - $($failures -join "`n - ")"
}

$verificationScope = if ($publishedRootPath) { 'layout templates and published pages' } else { 'layout templates' }
Write-Output "Content layout contracts verified against $verificationScope."
