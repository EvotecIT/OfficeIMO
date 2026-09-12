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
$licenseRoot = Join-Path $converterRoot 'licenses'
$managedAesNoticePath = Join-Path $licenseRoot 'OfficeIMO.Core-THIRD-PARTY-NOTICES.md'
$japaneseFontLicensePath = Join-Path $licenseRoot 'OFL-NotoCJK.txt'
$convertPagePath = Join-Path $SiteRoot 'convert/index.html'
$conversionGuidesPath = Join-Path $SiteRoot 'convert/guides/index.html'
$redirectManifestPath = Join-Path $SiteRoot '_powerforge/redirects.json'

foreach ($path in @(
        $indexPath,
        $modulePath,
        $appAssemblyPath,
        $runtimeWasmPath,
        $managedAesNoticePath,
        $japaneseFontLicensePath,
        $convertPagePath,
        $conversionGuidesPath,
        $redirectManifestPath
    )) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Converter publish is missing '$path'."
    }
}

$converterFramePattern = 'data-workspace-src="/apps/officeimo-converter/\?embedded=1"'
$convertPage = Get-Content -LiteralPath $convertPagePath -Raw
if ($convertPage -notmatch $converterFramePattern) {
    throw "The primary /convert/ route does not host the browser converter."
}

$conversionGuides = Get-Content -LiteralPath $conversionGuidesPath -Raw
if ($conversionGuides -notmatch '<h1>Document Conversion Guides for \.NET</h1>') {
    throw "The /convert/guides/ route does not contain the conversion guide."
}

$redirectManifest = Get-Content -LiteralPath $redirectManifestPath -Raw | ConvertFrom-Json
$playgroundRedirect = @($redirectManifest.redirects | Where-Object {
    $_.from -eq '/playground/' -and $_.to -eq '/convert/' -and
    $_.status -eq 301 -and $_.preserveQuery -eq $true
})
if ($playgroundRedirect.Count -ne 1) {
    throw 'The compatibility /playground/ route must redirect to /convert/ and preserve workflow query parameters.'
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
if ($module -notmatch 'export function createObjectUrl') {
    throw 'Converter collocated interop module is incomplete.'
}

& (Join-Path $PSScriptRoot 'Test-ConverterAssetGraph.ps1') -SiteRoot $converterRoot

Write-Output "Converter publish verified: $converterRoot ($([System.IO.Path]::GetFileName($appAssemblyPath)))"
