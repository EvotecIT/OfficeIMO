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

foreach ($path in @($indexPath, $modulePath, $appAssemblyPath, $runtimeWasmPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Converter publish is missing '$path'."
    }
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
if ($converterCss -notmatch '\.ocx-embedded \.ocx-appbar\s*\{\s*display:\s*none') {
    throw 'Converter stylesheet does not hide the standalone app bar in embedded mode.'
}

$module = Get-Content -LiteralPath $modulePath -Raw
if ($module -notmatch 'export function createObjectUrl' -or $module -notmatch 'export function revokeObjectUrl') {
    throw 'Converter collocated interop module is incomplete.'
}

Write-Output "Converter publish verified: $converterRoot ($([System.IO.Path]::GetFileName($appAssemblyPath)))"
