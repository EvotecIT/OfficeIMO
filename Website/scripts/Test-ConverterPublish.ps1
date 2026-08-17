[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $SiteRoot
)

$ErrorActionPreference = 'Stop'
$converterRoot = Join-Path $SiteRoot 'apps/officeimo-converter'
$indexPath = Join-Path $converterRoot 'index.html'
$modulePath = Join-Path $converterRoot 'Components/ConverterWorkspace.razor.js'
$frameworkRoot = Join-Path $converterRoot '_framework'
$appAssemblyPath = Get-ChildItem -LiteralPath $frameworkRoot -File -Filter 'OfficeIMO.Web.Converter*.wasm' -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notmatch '\.(br|gz)$' } |
    Select-Object -First 1 -ExpandProperty FullName
$runtimeWasmPath = Get-ChildItem -LiteralPath $frameworkRoot -File -Filter 'dotnet.native*.wasm' -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notmatch '\.(br|gz)$' } |
    Select-Object -First 1 -ExpandProperty FullName
$convertPagePath = Join-Path $SiteRoot 'convert/index.html'
$conversionGuidesPath = Join-Path $SiteRoot 'convert/guides/index.html'
$playgroundPagePath = Join-Path $SiteRoot 'playground/index.html'

foreach ($path in @(
        $indexPath,
        $modulePath,
        $appAssemblyPath,
        $runtimeWasmPath,
        $convertPagePath,
        $conversionGuidesPath,
        $playgroundPagePath
    )) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Converter publish is missing '$path'."
    }
}

$converterFramePattern = 'src="/apps/officeimo-converter/\?embedded=1"'
$convertPage = Get-Content -LiteralPath $convertPagePath -Raw
if ($convertPage -notmatch $converterFramePattern) {
    throw "The primary /convert/ route does not host the browser converter."
}

$conversionGuides = Get-Content -LiteralPath $conversionGuidesPath -Raw
if ($conversionGuides -notmatch '<h1>Document Conversion Guides for \.NET</h1>') {
    throw "The /convert/guides/ route does not contain the conversion guide."
}

$playgroundPage = Get-Content -LiteralPath $playgroundPagePath -Raw
if ($playgroundPage -notmatch $converterFramePattern) {
    throw "The compatibility /playground/ route does not host the browser converter."
}
$canonicalLink = [regex]::Matches(
    $playgroundPage,
    '<link\b[^>]*>',
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
) | Where-Object {
    $_.Value -match '\brel\s*=\s*(?:"canonical"|''canonical''|canonical)(?:\s|/?>)' -and
    $_.Value -match '\bhref\s*=\s*(?:"https://officeimo\.com/convert/"|''https://officeimo\.com/convert/''|https://officeimo\.com/convert/)(?:\s|/?>)'
} | Select-Object -First 1
if (-not $canonicalLink) {
    throw "The compatibility /playground/ route does not canonicalize to /convert/."
}
$robotsMeta = [regex]::Matches(
    $playgroundPage,
    '<meta\b[^>]*>',
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
) | Where-Object {
    $_.Value -match '\bname\s*=\s*(?:"robots"|''robots''|robots)(?:\s|/?>)' -and
    $_.Value -match '\bcontent\s*=\s*(?:"[^"]*\bnoindex\b[^"]*"|''[^'']*\bnoindex\b[^'']*''|[^\s>]*\bnoindex\b[^\s>]*)(?:\s|/?>)'
} | Select-Object -First 1
if (-not $robotsMeta) {
    throw "The compatibility /playground/ route is indexable instead of being a noindex alias."
}

$runtimeWasm = [System.Text.Encoding]::ASCII.GetString(
    [System.IO.File]::ReadAllBytes($runtimeWasmPath)
)
if ($runtimeWasm -notmatch 'hb_blob_create') {
    throw "Converter runtime '$runtimeWasmPath' does not contain the HarfBuzz native symbols required by the faithful PDF profile. Install the wasm-tools workload and publish with WasmBuildNative enabled."
}

$index = Get-Content -LiteralPath $indexPath -Raw
if ($index -notmatch '<base href="/apps/officeimo-converter/"') {
    throw 'Converter index does not use the production base path.'
}
if ($index -match 'converter-interop\.js') {
    throw 'Converter index still references the removed global interop script.'
}
if ($index -notmatch '_framework/blazor\.webassembly') {
    throw 'Converter index does not reference the Blazor WebAssembly bootstrap.'
}
if ($index -notmatch "embedded'\)===\'1\'" -or $index -notmatch "classList\.add\('ocx-embedded'\)") {
    throw 'Converter index does not enable the shared-shell embedded mode.'
}

$converterCssPath = Join-Path $converterRoot 'converter.css'
$converterCss = Get-Content -LiteralPath $converterCssPath -Raw
if ($converterCss -notmatch '\.ocx-embedded \.ocx-site-header' -or
    $converterCss -notmatch '\.ocx-embedded \.ocx-site-footer\s*\{\s*display:\s*none') {
    throw 'Converter stylesheet does not hide the standalone site shell in embedded mode.'
}
if ($converterCss -notmatch '\.ocx-hidden-input\s*\{[^}]*\binset:\s*0' -or
    $converterCss -match '\.ocx-hidden-input\s*\{[^}]*\bpointer-events:\s*none') {
    throw 'Converter file inputs do not cover the visible dropzone as native click targets.'
}

$module = Get-Content -LiteralPath $modulePath -Raw
if ($module -notmatch 'export function createObjectUrl' -or $module -notmatch 'export function revokeObjectUrl') {
    throw 'Converter collocated interop module is incomplete.'
}

& (Join-Path $PSScriptRoot 'Test-ConverterAssetGraph.ps1') -SiteRoot $converterRoot

Write-Output "Converter publish verified: $converterRoot ($([System.IO.Path]::GetFileName($appAssemblyPath)))"
