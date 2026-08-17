#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Verifies rendered customer-facing presentation contracts.
#>

[CmdletBinding()]
param(
    [string] $SiteRoot = (Join-Path (Split-Path -Parent $PSScriptRoot) '_site'),
    [string] $SourceRoot = (Split-Path -Parent $PSScriptRoot)
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
Assert-ContainsLiteral -Text $productCss -Expected '.imo-intent-content > article :is(ul, ol)' -Contract 'solution list presentation'
Assert-ContainsLiteral -Text $productCss -Expected '.imo-intent-content > article h2' -Contract 'prose divider scope'
Assert-ContainsLiteral -Text $productCss -Expected '.imo-capability-state[data-state="Native"]' -Contract 'capability state presentation'
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

foreach ($operation in @($pdfWorkflowCatalog.operations)) {
    $operationHtml = Get-RequiredText -Path (Join-Path $siteRootPath "pdf\$($operation.slug)\index.html")
    $expectedUrl = "$($pdfWorkflowCatalog.browserBaseUrl)$($operation.id)"
    Assert-ContainsLiteral -Text $operationHtml -Expected $expectedUrl -Contract "PDF operation handoff '$($operation.id)'"
    Assert-ContainsLiteral -Text $operationHtml -Expected '/css/product.css' -Contract "PDF operation styling '$($operation.id)'"
}

foreach ($conversion in @($pdfWorkflowCatalog.browserConversions)) {
    $conversionHtml = Get-RequiredText -Path (Join-Path $siteRootPath "convert\$($conversion.slug)\index.html")
    $expectedUrl = "/apps/officeimo-converter/?workspace=convert&route=$($conversion.routeId)"
    Assert-ContainsLiteral -Text $conversionHtml -Expected $expectedUrl -Contract "PDF conversion handoff '$($conversion.routeId)'"
}

Write-Host 'Website presentation contracts validated.'
