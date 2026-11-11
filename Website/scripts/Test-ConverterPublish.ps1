[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $SiteRoot
)

# Verifies the published browser tools: the static /browser/<tool>/ pages, the /convert/ directory,
# and the OfficeIMO engine that the pages run in a Web Worker from /apps/officeimo-converter/.

$ErrorActionPreference = 'Stop'
$converterRoot = Join-Path $SiteRoot 'apps/officeimo-converter'
$indexPath = Join-Path $converterRoot 'index.html'
$workerPath = Join-Path $converterRoot 'engine-worker.js'
$siteScriptPath = Join-Path $SiteRoot 'js/site.js'
$frameworkRoot = Join-Path $converterRoot '_framework'
$dotnetPath = Join-Path $frameworkRoot 'dotnet.js'
$appAssemblyPath = Get-ChildItem -LiteralPath $frameworkRoot -File -Filter 'OfficeIMO.Web.Converter*.wasm' -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notmatch '\.(br|gz)$' } |
    Select-Object -First 1 -ExpandProperty FullName
$runtimeWasmPath = Get-ChildItem -LiteralPath $frameworkRoot -File -Filter 'dotnet.native*.wasm' -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notmatch '\.(br|gz)$' } |
    Select-Object -First 1 -ExpandProperty FullName
$licenseRoot = Join-Path $converterRoot 'licenses'
$managedAesNoticePath = Join-Path $licenseRoot 'OfficeIMO.Core-THIRD-PARTY-NOTICES.md'
$japaneseFontLicensePath = Join-Path $licenseRoot 'OFL-NotoCJK.txt'
$convertPagePath = Join-Path $SiteRoot 'convert/index.html'
$conversionGuidesPath = Join-Path $SiteRoot 'convert/guides/index.html'
$sharingGuidePath = Join-Path $SiteRoot 'docs/before-sharing/index.html'
$redirectManifestPath = Join-Path $SiteRoot '_powerforge/redirects.json'

foreach ($path in @(
        $indexPath,
        $workerPath,
        $dotnetPath,
        $siteScriptPath,
        $appAssemblyPath,
        $runtimeWasmPath,
        $managedAesNoticePath,
        $japaneseFontLicensePath,
        $convertPagePath,
        $conversionGuidesPath,
        $sharingGuidePath,
        $redirectManifestPath
    )) {
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Browser tools publish is missing '$path'."
    }
}

$catalog = Get-Content (Join-Path $PSScriptRoot '../data/browser_tools.json') -Raw | ConvertFrom-Json
$toolScriptPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
$convertPage = Get-Content -LiteralPath $convertPagePath -Raw
if ($convertPage -match '<iframe\b' -or $convertPage -match 'browser-workspace-template') {
    throw 'The /convert/ directory must be plain HTML; tools open on their own /browser/ pages.'
}
foreach ($tool in $catalog.tools) {
    if ($convertPage -notmatch ('\bdata-tool=["'']?' + [regex]::Escape($tool.id) + '["'']?(?=\s|>)') -or
        $convertPage -notmatch ('href=["'']?/browser/' + [regex]::Escape($tool.id) + '/["'']?')) {
        throw "The /convert/ directory is missing a card for '$($tool.id)'."
    }
    $toolPagePath = Join-Path $SiteRoot "browser/$($tool.id)/index.html"
    if (-not (Test-Path -LiteralPath $toolPagePath -PathType Leaf)) {
        throw "Browser tool page '/browser/$($tool.id)/' was not generated. Run scripts/Sync-BrowserToolPages.ps1."
    }
    $toolPage = Get-Content -LiteralPath $toolPagePath -Raw
    $toolScriptReference = [regex]::Match($toolPage, '<script\b[^>]*\ssrc=["'']?(?<path>/js/browser-tool(?:\.[a-f0-9]+)?\.js)["'']?(?=\s|/?>)')
    if ($toolPage -notmatch ('\bdata-browser-tool=["'']?' + [regex]::Escape($tool.id) + '["'']?(?=\s|>)') -or
        $toolPage -notmatch 'data-engine-base=["'']?/apps/officeimo-converter/' -or
        $toolPage -notmatch '<link[^>]+href=["'']?/css/browser-tool(?:\.[a-f0-9]+)?\.css' -or
        -not $toolScriptReference.Success) {
        throw "Browser tool page '/browser/$($tool.id)/' is missing its tool wiring, stylesheet, or script."
    }
    $referencedScriptPath = Join-Path $SiteRoot $toolScriptReference.Groups['path'].Value.TrimStart('/')
    if (-not (Test-Path -LiteralPath $referencedScriptPath -PathType Leaf)) {
        throw "Browser tool page '/browser/$($tool.id)/' references missing script '$referencedScriptPath'."
    }
    $null = $toolScriptPaths.Add($referencedScriptPath)
    if ($toolPage -notmatch [regex]::Escape([System.Net.WebUtility]::HtmlEncode($tool.title))) {
        throw "Browser tool page '/browser/$($tool.id)/' does not show its title."
    }
}

$conversionGuides = Get-Content -LiteralPath $conversionGuidesPath -Raw
if ($conversionGuides -notmatch '<h1>Document Conversion Guides for \.NET</h1>') {
    throw "The /convert/guides/ route does not contain the conversion guide."
}

$sharingGuide = Get-Content -LiteralPath $sharingGuidePath -Raw
if ($sharingGuide -notmatch '<h1(?:\s[^>]*)?>Prepare a document for sharing</h1>' -or
    $sharingGuide -notmatch 'Content Credentials' -or
    $sharingGuide -notmatch 'href=["'']?/browser/clean-before-sharing/["'']?(?:\s|/?>)' -or
    $sharingGuide -notmatch 'href=["'']?/browser/hidden-characters/["'']?(?:\s|/?>)' -or
    $sharingGuide -notmatch 'href=["'']?https://github\.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo\.hidden-data-support-matrix\.md["'']?(?:\s|/?>)' -or
    $sharingGuide -notmatch 'href=["'']?https://github\.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo\.provenance-support-matrix\.md["'']?(?:\s|/?>)') {
    throw 'The /docs/before-sharing/ guide does not explain the clean-before-sharing workflow and link its tools and support matrices.'
}

$redirectManifest = Get-Content -LiteralPath $redirectManifestPath -Raw | ConvertFrom-Json
$playgroundRedirect = @($redirectManifest.redirects | Where-Object {
    $_.from -eq '/playground/' -and $_.to -eq '/convert/' -and
    $_.status -eq 301 -and $_.preserveQuery -eq $true
})
if ($playgroundRedirect.Count -ne 1) {
    throw 'The compatibility /playground/ route must redirect to /convert/ and preserve workflow query parameters.'
}
foreach ($retired in @(
    @{ From = '/browser/file-origin/'; To = '/browser/clean-before-sharing/' },
    @{ From = '/provenance/'; To = '/docs/before-sharing/provenance/' })) {
    if (@($redirectManifest.redirects | Where-Object { $_.from -eq $retired.From -and $_.to -eq $retired.To -and $_.status -eq 301 }).Count -ne 1) {
        throw "The retired $($retired.From) route must redirect to $($retired.To)."
    }
}

# Older links (/convert/?route=docx-pdf) resolve through the directory's legacy keys.
$siteScript = Get-Content -LiteralPath $siteScriptPath -Raw
if ($siteScript -notmatch 'data-legacy' -or $siteScript -notmatch 'location\.replace') {
    throw 'The /convert/ directory script no longer forwards legacy workspace links to tool pages.'
}
foreach ($tool in $catalog.tools) {
    foreach ($legacy in @($tool.legacy)) {
        if ($convertPage -notmatch [regex]::Escape([System.Net.WebUtility]::HtmlEncode($legacy))) {
            throw "Legacy link '$legacy' for '$($tool.id)' is not declared on the /convert/ directory."
        }
    }
}

$runtimeWasm = [System.Text.Encoding]::ASCII.GetString(
    [System.IO.File]::ReadAllBytes($runtimeWasmPath)
)
if ($runtimeWasm -notmatch 'hb_blob_create') {
    throw "Engine runtime '$runtimeWasmPath' does not contain the HarfBuzz native symbols required by the faithful PDF profile. Install the wasm-tools workload and publish with WasmBuildNative enabled."
}

$index = Get-Content -LiteralPath $indexPath -Raw
if ($index -notmatch 'location\.replace\(\s*["'']/convert/["'']\s*\+\s*location\.search\s*\)' -or $index -match '_framework/blazor') {
    throw 'The engine folder index must only forward older app links to the /convert/ directory.'
}

$worker = Get-Content -LiteralPath $workerPath -Raw
if ($worker -notmatch 'from "\./_framework/dotnet\.js"' -or
    $worker -notmatch 'loadLazyAssembly' -or
    $worker -notmatch 'getAssemblyExports') {
    throw 'The engine worker must boot ./_framework/dotnet.js, load engines lazily, and call the exported tool entry points.'
}

# Verify the actual assets loaded by the pages, including fingerprinted copies.
foreach ($referencedScriptPath in $toolScriptPaths) {
    $toolScript = Get-Content -LiteralPath $referencedScriptPath -Raw
    # Minification renames the local configuration variable and normalizes quotes and spacing.
    if ($toolScript -cnotmatch 'new\s+Worker\(\s*[$A-Za-z_][$\w]*\.base\s*\+\s*["'']engine-worker\.js["'']\s*,\s*\{\s*type\s*:\s*["'']module["'']\s*\}\s*\)') {
        throw "Tool page script '$referencedScriptPath' must run the engine in a module Web Worker."
    }
}

& (Join-Path $PSScriptRoot 'Test-ConverterAssetGraph.ps1') -SiteRoot $converterRoot

Write-Output "Browser tools publish verified: $($catalog.tools.Count) tool pages, engine $([System.IO.Path]::GetFileName($appAssemblyPath))"
