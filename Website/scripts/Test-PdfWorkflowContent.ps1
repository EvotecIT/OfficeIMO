#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Verifies PDF workflow content against the browser and conversion catalogs.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$siteRootPath = (Resolve-Path -LiteralPath $SiteRoot).Path

function Get-FrontMatterValue {
    param(
        [Parameter(Mandatory)] [string] $Content,
        [Parameter(Mandatory)] [string] $Name
    )

    $pattern = '(?m)^{0}:\s*["'']?(?<value>[^"''\r\n]+)["'']?\s*$' -f [regex]::Escape($Name)
    $match = [regex]::Match($Content, $pattern)
    if (-not $match.Success) {
        throw "Required front-matter value '$Name' is missing."
    }
    return $match.Groups['value'].Value.Trim()
}

function Assert-WorkflowPage {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Identity,
        [Parameter(Mandatory)] [string] $PrimaryUrl
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "PDF workflow page is missing: $Path"
    }

    $content = Get-Content -LiteralPath $Path -Raw
    $actualIdentity = Get-FrontMatterValue -Content $content -Name 'meta.workflow_id'
    if ($actualIdentity -ne $Identity) {
        throw "PDF workflow page '$Path' identifies '$actualIdentity'; expected '$Identity'."
    }

    $actualPrimaryUrl = Get-FrontMatterValue -Content $content -Name 'meta.primary_url'
    if ($actualPrimaryUrl -ne $PrimaryUrl) {
        throw "PDF workflow page '$Path' links to '$actualPrimaryUrl'; expected '$PrimaryUrl'."
    }

    $description = Get-FrontMatterValue -Content $content -Name 'description'
    if ($description.Length -lt 120 -or $description.Length -gt 160) {
        throw "PDF workflow page '$Path' has a $($description.Length)-character description; expected 120-160."
    }

    $stepCount = [regex]::Matches($content, '(?m)^\s{2}- name:\s*').Count
    if ($stepCount -lt 3) {
        throw "PDF workflow page '$Path' must expose at least three visible HowTo steps."
    }
    if (-not $content.Contains('```csharp', [StringComparison]::Ordinal)) {
        throw "PDF workflow page '$Path' must include a public C# example."
    }
}

$catalogPath = Join-Path $siteRootPath 'data\pdf_workflows.json'
$catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json
if ([int] $catalog.schemaVersion -ne 1) {
    throw 'PDF workflow catalog schemaVersion must be 1.'
}

$operations = @($catalog.operations)
if ($operations.Count -ne 12) {
    throw "PDF workflow catalog must contain the 12 browser operations; found $($operations.Count)."
}
if (@($operations.id | Sort-Object -Unique).Count -ne $operations.Count -or
    @($operations.slug | Sort-Object -Unique).Count -ne $operations.Count) {
    throw 'PDF workflow ids and slugs must be unique.'
}

$browserCatalogPath = Join-Path $siteRootPath 'Apps\OfficeIMO.Web.Converter\Services\PdfToolCatalog.cs'
$browserCatalogText = Get-Content -LiteralPath $browserCatalogPath -Raw
$browserIds = @([regex]::Matches($browserCatalogText, 'new\("(?<id>[a-z0-9-]+)"') |
    ForEach-Object { $_.Groups['id'].Value })
$contentIds = @($operations.id)
$missingContent = @($browserIds | Where-Object { $_ -notin $contentIds })
$missingBrowserTools = @($contentIds | Where-Object { $_ -notin $browserIds })
if ($missingContent.Count -gt 0 -or $missingBrowserTools.Count -gt 0 -or $browserIds.Count -ne $contentIds.Count) {
    throw "PDF browser tools and public workflow catalog differ. Missing content: $($missingContent -join ', '). Missing browser tools: $($missingBrowserTools -join ', ')."
}

foreach ($operation in $operations) {
    $pagePath = Join-Path $siteRootPath "content\pdf-workflows\$($operation.slug).md"
    $browserUrl = [string] $catalog.browserBaseUrl + [string] $operation.id
    Assert-WorkflowPage -Path $pagePath -Identity ([string] $operation.id) -PrimaryUrl $browserUrl
}

$reorder = @($operations | Where-Object id -eq 'reorder')
$reorderPageText = Get-Content -LiteralPath (Join-Path $siteRootPath 'content\pdf-workflows\reorder-pages.md') -Raw
if ($reorder.Count -ne 1 -or
    -not ([string] $reorder[0].summary).Contains('every source page exactly once', [StringComparison]::OrdinalIgnoreCase) -or
    -not $reorderPageText.Contains('every source page exactly once', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The public reorder workflow must state the full-permutation contract: every source page exactly once.'
}

$conversionCatalog = Get-Content -LiteralPath (Join-Path $siteRootPath 'data\office_conversion_routes.json') -Raw | ConvertFrom-Json
$browserConversions = @($catalog.browserConversions)
if ($browserConversions.Count -ne 4) {
    throw "PDF workflow catalog must expose the four browser PDF imports; found $($browserConversions.Count)."
}
foreach ($item in $browserConversions) {
    $routes = @($conversionCatalog.routes | Where-Object id -eq $item.routeId)
    if ($routes.Count -ne 1 -or [string] $routes[0].source -ne 'PDF' -or -not [bool] $routes[0].browserAvailable) {
        throw "PDF browser conversion '$($item.routeId)' is missing, duplicated, not PDF-sourced, or unavailable in the browser catalog."
    }
    $expectedTitle = "$($routes[0].source) to $($routes[0].target)"
    if ([string] $item.title -ne $expectedTitle -or [string] $item.summary -ne [string] $routes[0].description) {
        throw "PDF browser conversion '$($item.routeId)' presentation differs from the conversion catalog."
    }
    $pagePath = Join-Path $siteRootPath "content\conversions\$($item.slug).md"
    $browserUrl = "/apps/officeimo-converter/?workspace=convert&route=$($item.routeId)"
    Assert-WorkflowPage -Path $pagePath -Identity ([string] $item.routeId) -PrimaryUrl $browserUrl
}

$hubPath = Join-Path $siteRootPath 'content\pdf-workflows\index.md'
if (-not (Test-Path -LiteralPath $hubPath -PathType Leaf)) {
    throw "PDF workflow hub is missing: $hubPath"
}
$hubText = Get-Content -LiteralPath $hubPath -Raw
if (-not $hubText.Contains('{{< pdf-workflows >}}', [StringComparison]::Ordinal)) {
    throw 'PDF workflow hub must render the catalog-backed workflow index.'
}

Write-Host "PDF workflow content verified against $($operations.Count) browser tools and $($browserConversions.Count) browser conversions."
