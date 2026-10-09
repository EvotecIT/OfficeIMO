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
$toolScriptPath = Join-Path $SiteRoot 'js/browser-tool.js'
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
$provenanceGuidePath = Join-Path $SiteRoot 'provenance/index.html'
$redirectManifestPath = Join-Path $SiteRoot '_powerforge/redirects.json'

foreach ($path in @(
        $indexPath,
        $workerPath,
        $dotnetPath,
        $siteScriptPath,
        $toolScriptPath,
        $appAssemblyPath,
        $runtimeWasmPath,
        $managedAesNoticePath,
        $japaneseFontLicensePath,
        $convertPagePath,
        $conversionGuidesPath,
        $provenanceGuidePath,
        $redirectManifestPath
    )) {
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Browser tools publish is missing '$path'."
    }
}

$catalog = Get-Content (Join-Path $PSScriptRoot '../data/browser_tools.json') -Raw | ConvertFrom-Json
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
    if ($toolPage -notmatch ('\bdata-browser-tool=["'']?' + [regex]::Escape($tool.id) + '["'']?(?=\s|>)') -or
        $toolPage -notmatch 'data-engine-base=["'']?/apps/officeimo-converter/' -or
        $toolPage -notmatch '<link[^>]+href=["'']?/css/browser-tool(?:\.[a-f0-9]+)?\.css' -or
        $toolPage -notmatch '<script[^>]+src=["'']?/js/browser-tool(?:\.[a-f0-9]+)?\.js') {
        throw "Browser tool page '/browser/$($tool.id)/' is missing its tool wiring, stylesheet, or script."
    }
    if ($toolPage -notmatch [regex]::Escape([System.Net.WebUtility]::HtmlEncode($tool.title))) {
        throw "Browser tool page '/browser/$($tool.id)/' does not show its title."
    }
}

$conversionGuides = Get-Content -LiteralPath $conversionGuidesPath -Raw
if ($conversionGuides -notmatch '<h1>Document Conversion Guides for \.NET</h1>') {
    throw "The /convert/guides/ route does not contain the conversion guide."
}

$provenanceGuide = Get-Content -LiteralPath $provenanceGuidePath -Raw
if ($provenanceGuide -notmatch '<h1(?:\s[^>]*)?>Check and remove file provenance</h1>' -or
    $provenanceGuide -notmatch '<body class="imo-body imo-body--docs imo-body--conversion">' -or
    $provenanceGuide -notmatch '<link[^>]+href=["'']?/css/product(?:\.[a-f0-9]+)?\.css["'']?(?:\s|/?>)' -or
    $provenanceGuide -notmatch '<link[^>]+href=["'']?/css/docs(?:\.[a-f0-9]+)?\.css["'']?(?:\s|/?>)' -or
    $provenanceGuide -notmatch '<h2>Inspection summary</h2>' -or
    $provenanceGuide -notmatch 'OfficeIMO\.Workflows \+ format packages' -or
    $provenanceGuide -notmatch 'href=["'']?https://www\.nuget\.org/packages/OfficeIMO\.Workflows["'']?(?:\s|/?>)' -or
    $provenanceGuide -notmatch 'Content Credentials' -or
    $provenanceGuide -notmatch 'does not remove visible watermarks' -or
    $provenanceGuide -notmatch 'href="/browser/file-origin/"' -or
    $provenanceGuide -notmatch 'href="/browser/hidden-characters/"' -or
    $provenanceGuide -notmatch 'href=["'']?https://github\.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo\.provenance-support-matrix\.md["'']?(?:\s|/?>)') {
    throw 'The /provenance/ route does not explain the supported file-origin workflow and its limits.'
}

$redirectManifest = Get-Content -LiteralPath $redirectManifestPath -Raw | ConvertFrom-Json
$playgroundRedirect = @($redirectManifest.redirects | Where-Object {
    $_.from -eq '/playground/' -and $_.to -eq '/convert/' -and
    $_.status -eq 301 -and $_.preserveQuery -eq $true
})
if ($playgroundRedirect.Count -ne 1) {
    throw 'The compatibility /playground/ route must redirect to /convert/ and preserve workflow query parameters.'
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
if ($index -notmatch "location\.replace\('/convert/' \+ location\.search\)" -or $index -match '_framework/blazor') {
    throw 'The engine folder index must only forward older app links to the /convert/ directory.'
}

$worker = Get-Content -LiteralPath $workerPath -Raw
if ($worker -notmatch 'from "\./_framework/dotnet\.js"' -or
    $worker -notmatch 'loadLazyAssembly' -or
    $worker -notmatch 'getAssemblyExports') {
    throw 'The engine worker must boot ./_framework/dotnet.js, load engines lazily, and call the exported tool entry points.'
}

$toolScript = Get-Content -LiteralPath $toolScriptPath -Raw
if ($toolScript -notmatch "new Worker\(cfg\.base \+ 'engine-worker\.js', \{ type: 'module' \}\)") {
    throw 'The tool page script must run the engine in a module Web Worker.'
}

& (Join-Path $PSScriptRoot 'Test-ConverterAssetGraph.ps1') -SiteRoot $converterRoot

Write-Output "Browser tools publish verified: $($catalog.tools.Count) tool pages, engine $([System.IO.Path]::GetFileName($appAssemblyPath))"
