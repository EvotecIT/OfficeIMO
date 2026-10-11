#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Verifies rendered customer-facing presentation contracts.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) '_site'),
    [string] $SourceRoot = (Split-Path -Parent $PSScriptRoot),
    [ValidateSet('All', 'Showcase')][string] $Scope = 'All'
)

$ErrorActionPreference = 'Stop'

function Get-RequiredText {
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required presentation artifact is missing: $Path"
    }

    return [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $Path).Path)
}

function Assert-ContainsLiteral {
    param(
        [Parameter(Mandatory)]
        [string] $Text,

        [Parameter(Mandatory)]
        [string] $Expected,

        [Parameter(Mandatory)]
        [string] $Contract
    )

    if (-not $Text.Contains($Expected)) {
        throw "Presentation contract '$Contract' is missing '$Expected'."
    }
}

$siteRootPath = (Resolve-Path -LiteralPath $SiteRoot).Path
$sourceRootPath = (Resolve-Path -LiteralPath $SourceRoot).Path

function Assert-ContainsAttribute {
    param(
        [Parameter(Mandatory)] [string] $Text,
        [Parameter(Mandatory)] [string] $Name,
        [Parameter(Mandatory)] [string] $Value,
        [Parameter(Mandatory)] [string] $Contract
    )

    $escapedValue = [regex]::Escape($Value)
    $pattern = '\s' + [regex]::Escape($Name) + '\s*=\s*(?:"' + $escapedValue + '"|''' + $escapedValue + '''|' + $escapedValue + '(?=\s|/?>))'
    if ($Text -notmatch $pattern) {
        throw "Presentation contract '$Contract' is missing attribute '$Name=$Value'."
    }
}

function Assert-ContainsPattern {
    param(
        [Parameter(Mandatory)] [string] $Text,
        [Parameter(Mandatory)] [string] $Pattern,
        [Parameter(Mandatory)] [string] $Contract
    )

    if ($Text -cnotmatch $Pattern) {
        throw "Presentation contract '$Contract' is missing."
    }
}

$showcase = Get-Content -LiteralPath (Join-Path $sourceRootPath 'data/showcase.json') -Raw | ConvertFrom-Json -Depth 40
$galleryHtml = Get-RequiredText -Path (Join-Path $siteRootPath 'showcase/index.html')
Assert-ContainsAttribute -Text $galleryHtml -Name 'id' -Value 'showcase-language' -Contract 'showcase language selection'
foreach ($card in @($showcase.cards | Where-Object walkthrough_url)) {
    $cardId = [regex]::Escape($card.id)
    $cardIdValue = '(?:"' + $cardId + '"|''' + $cardId + '''|' + $cardId + '(?=\s|/?>))'
    $galleryCard = [regex]::Match($galleryHtml, '<article\b(?=[^>]*\sid\s*=\s*' + $cardIdValue + ')[^>]*>').Value
    if ([string]::IsNullOrWhiteSpace($galleryCard)) {
        throw "Showcase gallery is missing walkthrough card '$($card.id)'."
    }
    Assert-ContainsPattern -Text $galleryCard -Pattern 'data-showcase-search="[^"]*\bwalkthrough\b' -Contract "$($card.id) walkthrough discovery"
    $languageLabel = ($showcase.source_languages | Where-Object id -CEQ $card.source_language).label
    $walkthroughHtml = Get-RequiredText -Path (Join-Path $siteRootPath "showcase/$($card.id)/index.html")
    Assert-ContainsLiteral -Text $walkthroughHtml -Expected "$languageLabel walkthrough" -Contract "$($card.id) source language"
    Assert-ContainsLiteral -Text $walkthroughHtml -Expected "Copy $languageLabel source" -Contract "$($card.id) code copy label"
    Assert-ContainsAttribute -Text $walkthroughHtml -Name 'class' -Value "language-$($card.source_language)" -Contract "$($card.id) syntax highlighting"
    if ($card.source_language -eq 'powershell') {
        $moduleVersion = [string](Import-PowerShellDataFile -LiteralPath (Join-Path $sourceRootPath 'data/apidocs/powershell/PSWriteOffice.psd1')).ModuleVersion
        $runCommand = $card.run_command.Replace('{moduleVersion}', $moduleVersion)
        Assert-ContainsLiteral -Text $walkthroughHtml -Expected ([Net.WebUtility]::HtmlEncode($runCommand)) -Contract "$($card.id) executable command"
        Assert-ContainsLiteral -Text $walkthroughHtml -Expected ([Net.WebUtility]::HtmlEncode($runCommand.Replace('pwsh -NoProfile', 'powershell -NoProfile'))) -Contract "$($card.id) Windows PowerShell command"
        $commandMetadata = Get-Content -LiteralPath (Join-Path $sourceRootPath 'data/apidocs/powershell/command-metadata.json') -Raw | ConvertFrom-Json
        $aliases = @{}
        foreach ($command in $commandMetadata.commands) {
            foreach ($alias in $command.aliases) { $aliases[$alias] = $command.name }
        }
        $apiLinks = Get-Content -LiteralPath (Join-Path $sourceRootPath 'data/api_links.json') -Raw | ConvertFrom-Json
        foreach ($link in $apiLinks.pages.($card.walkthrough_url)) {
            if ($aliases.ContainsKey($link.n) -and $link.h -cne "/api/powershell/$($aliases[$link.n].ToLowerInvariant())/") {
                throw "PowerShell alias '$($link.n)' must link to its command reference on '$($card.id)'."
            }
        }
    }
}

if ($Scope -eq 'Showcase') { Write-Host 'Showcase presentation contracts validated.'; return }

$presentationRoutes = @('', 'downloads', 'studio', 'tool', 'products/excel', 'products/reader', 'libraries', 'convert', 'convert/guides', 'browser/word-to-pdf', 'pdf', 'pdf/merge', 'docs', 'api/word')
# Parse actual head elements, excluding comments, script text and inert fallbacks.
# Python's standard HTML parser also accepts the production optimizer's markup.
$pythonCommand = Get-Command python3, python -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $pythonCommand) { throw 'Python is required to verify first-paint styles.' }
$presentationFiles = @($presentationRoutes | ForEach-Object { Join-Path $siteRootPath "$_/index.html" })
& $pythonCommand.Source (Join-Path $PSScriptRoot 'Test-FirstPaintStyles.py') @presentationFiles
if ($LASTEXITCODE -ne 0) { throw 'First-paint stylesheet verification failed.' }

foreach ($route in $presentationRoutes) {
    $routeHtml = Get-RequiredText -Path (Join-Path $siteRootPath "$route/index.html")
    # The logo supplies Home; the main-menu-only verifier warning is baselined.
    if ($routeHtml -notmatch '<a\b(?=[^>]*\sclass=(?:"(?:[^"]*\s)?imo-header__logo(?:\s[^"]*)?"|''(?:[^'']*\s)?imo-header__logo(?:\s[^'']*)?''|imo-header__logo(?=\s|>)))(?=[^>]*\shref=["'']?/["'']?(?=\s|>))[^>]*>') {
        throw "Route '/$route/' must expose Home through the linked site logo."
    }
    $navigationCount = [regex]::Matches($routeHtml, '<nav\b[^>]*\sid=["'']?main-navigation["'']?(?=\s|>)').Count
    if ($navigationCount -ne 1) {
        throw "Route '/$route/' must render one global navigation menu; found $navigationCount."
    }
    foreach ($destination in @('/studio/', '/convert/', '/tool/', '/libraries/', '/docs/', '/downloads/')) {
        Assert-ContainsAttribute -Text $routeHtml -Name 'href' -Value $destination -Contract "global navigation on /$route/"
    }
}
foreach ($route in @('docs', 'api/word', 'convert/guides', 'convert/doc-docx', 'pdf', 'pdf/merge')) {
    $routeHtml = Get-RequiredText -Path (Join-Path $siteRootPath "$route/index.html")
    Assert-ContainsAttribute -Text $routeHtml -Name 'id' -Value 'documentation-navigation' -Contract "documentation navigation on /$route/"
    Assert-ContainsLiteral -Text $routeHtml -Expected 'imo-documentation-toolbar' -Contract "documentation toolbar on /$route/"
}

$solutionHtml = Get-RequiredText -Path (Join-Path $siteRootPath 'solutions\legacy-office-modernization\index.html')
$conversionHtml = Get-RequiredText -Path (Join-Path $siteRootPath 'convert\doc-docx\index.html')
$comparisonHtml = Get-RequiredText -Path (Join-Path $siteRootPath 'comparisons\officeimo-vs-closedxml-epplus\index.html')
$compatibilityHtml = Get-RequiredText -Path (Join-Path $siteRootPath 'compatibility\index.html')
$productCss = Get-RequiredText -Path (Join-Path $siteRootPath 'css\product.css')
$pdfWorkflowCatalogPath = Join-Path $sourceRootPath 'data\pdf_workflows.json'
$pdfWorkflowCatalog = Get-Content -LiteralPath $pdfWorkflowCatalogPath -Raw | ConvertFrom-Json -Depth 20

Assert-ContainsLiteral -Text $solutionHtml -Expected 'imo-intent-content imo-prose markdown-body' -Contract 'solution prose styling'
Assert-ContainsLiteral -Text $conversionHtml -Expected 'imo-intent-content imo-prose markdown-body' -Contract 'conversion prose styling'
Assert-ContainsLiteral -Text $comparisonHtml -Expected 'imo-comparison-detail-content imo-prose markdown-body' -Contract 'comparison prose styling'
Assert-ContainsPattern -Text $productCss -Pattern '\.imo-intent-content\s*>\s*article\s+:is\(ul,\s*ol\)' -Contract 'solution list presentation'
Assert-ContainsPattern -Text $productCss -Pattern '\.imo-intent-content\s*>\s*article\s+h2(?=\s|[,\{])' -Contract 'prose divider scope'
Assert-ContainsPattern -Text $productCss -Pattern '\.imo-capability-state\[data-state=["'']?Native["'']?\]' -Contract 'capability state presentation'
Assert-ContainsLiteral -Text $productCss -Expected '.imo-capability-card__source' -Contract 'source-first compatibility metadata'
$appCss = Get-RequiredText -Path (Join-Path $siteRootPath 'css\app.css')
Assert-ContainsLiteral -Text $appCss -Expected 'var(--imo-on-accent,#fff)' -Contract 'accent control contrast'

Assert-ContainsLiteral -Text $compatibilityHtml -Expected 'Current source contract' -Contract 'source-first compatibility label'
if ($compatibilityHtml -match 'imo-capability-card__version' -or
    $compatibilityHtml -match 'Package\s+\d+\.\d+\.\d+' -or
    $compatibilityHtml -match 'Package line') {
    throw 'Compatibility presentation still binds the current source contract to a fixed package version.'
}

$pdfHubHtml = Get-RequiredText -Path (Join-Path $siteRootPath 'pdf\index.html')
Assert-ContainsLiteral -Text $pdfHubHtml -Expected 'PDF tools for browsers and .NET' -Contract 'PDF workflow hub'

$browserTools = (Get-Content -LiteralPath (Join-Path $sourceRootPath 'data\browser_tools.json') -Raw | ConvertFrom-Json).tools
foreach ($operation in @($pdfWorkflowCatalog.operations)) {
    $operationHtml = Get-RequiredText -Path (Join-Path $siteRootPath "pdf\$($operation.slug)\index.html")
    $expectedUrl = "/browser/$((@($browserTools | Where-Object { $_.engine.kind -eq 'pdf' -and $_.engine.target -eq $operation.id }))[0].id)/"
    Assert-ContainsLiteral -Text $operationHtml -Expected $expectedUrl -Contract "PDF operation handoff '$($operation.id)'"
    Assert-ContainsLiteral -Text $operationHtml -Expected '/css/product.css' -Contract "PDF operation styling '$($operation.id)'"
}

foreach ($conversion in @($pdfWorkflowCatalog.browserConversions)) {
    $conversionHtml = Get-RequiredText -Path (Join-Path $siteRootPath "convert\$($conversion.slug)\index.html")
    $expectedUrl = "/browser/$((@($browserTools | Where-Object { $_.engine.kind -eq 'convert' -and $_.engine.target -eq $conversion.routeId }))[0].id)/"
    Assert-ContainsLiteral -Text $conversionHtml -Expected $expectedUrl -Contract "PDF conversion handoff '$($conversion.routeId)'"
}

Write-Host 'Website presentation contracts validated.'
