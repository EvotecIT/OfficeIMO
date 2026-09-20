[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $SiteRoot
)

$ErrorActionPreference = 'Stop'
$converterRoot = Join-Path $SiteRoot 'apps/officeimo-converter'
$indexPath = Join-Path $converterRoot 'index.html'
$modulePath = Join-Path $converterRoot 'Components/ConverterWorkspace.razor.js'
$workspaceModulePath = Join-Path $converterRoot 'Components/DocumentWorkspace.razor.js'
$siteScriptPath = Join-Path $SiteRoot 'js/site.js'
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
$provenanceGuidePath = Join-Path $SiteRoot 'provenance/index.html'
$redirectManifestPath = Join-Path $SiteRoot '_powerforge/redirects.json'

foreach ($path in @(
        $indexPath,
        $modulePath,
        $workspaceModulePath,
        $siteScriptPath,
        $appAssemblyPath,
        $runtimeWasmPath,
        $managedAesNoticePath,
        $japaneseFontLicensePath,
        $convertPagePath,
        $conversionGuidesPath,
        $provenanceGuidePath,
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
$withoutTemplates = [regex]::Replace($convertPage, '(?is)<template\b[^>]*>.*?</template>', '')
if ($withoutTemplates -match '<iframe\b' -or $convertPage -notmatch 'id="browser-workspace-template"') {
    throw 'The tool directory must keep its workspace iframe inert until a tool is selected.'
}
$routeCatalog = Get-Content (Join-Path $PSScriptRoot '../data/office_conversion_routes.json') -Raw | ConvertFrom-Json
$pdfCatalog = Get-Content (Join-Path $PSScriptRoot '../data/pdf_workflows.json') -Raw | ConvertFrom-Json
foreach ($route in @($routeCatalog.routes | Where-Object browserAvailable)) {
    if ($convertPage -notmatch ('data-route="' + [regex]::Escape($route.id) + '"')) {
        throw "The static directory is missing browser conversion '$($route.id)'."
    }
}
foreach ($tool in $pdfCatalog.operations) {
    if ($convertPage -notmatch ('data-pdf-tool="' + [regex]::Escape($tool.id) + '"')) {
        throw "The static directory is missing PDF tool '$($tool.id)'."
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
    $provenanceGuide -notmatch 'href="/convert/\?workspace=provenance"' -or
    $provenanceGuide -notmatch 'href=["'']?https://github\.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo\.provenance-support-matrix\.md["'']?(?:\s|/?>)' -or
    $provenanceGuide -notmatch 'carrier categories') {
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

$workspaceModule = Get-Content -LiteralPath $workspaceModulePath -Raw
$siteScript = Get-Content -LiteralPath $siteScriptPath -Raw
$titleValidation = [regex]::Match(
    $siteScript,
    'typeof\s+(?<selection>[A-Za-z_$][\w$]*)\.title\s*(?:!==|!=)\s*["'']string["'']'
)
$validatedSelection = if ($titleValidation.Success) {
    [regex]::Escape($titleValidation.Groups['selection'].Value)
}
if ($workspaceModule -notmatch 'officeimo:workspace-selection[^\r\n]+title' -or
    -not $titleValidation.Success -or
    $siteScript -notmatch "document\.title\s*=\s*$validatedSelection\.title") {
    throw 'The converter selection protocol does not propagate validated tool titles to the host page.'
}

& (Join-Path $PSScriptRoot 'Test-ConverterAssetGraph.ps1') -SiteRoot $converterRoot

Write-Output "Converter publish verified: $converterRoot ($([System.IO.Path]::GetFileName($appAssemblyPath)))"
