[CmdletBinding()]
param([string] $WebsiteRoot = (Split-Path -Parent $PSScriptRoot))

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ShowcaseEvidence.Helpers.ps1')
$manifest = Get-Content -LiteralPath (Join-Path $WebsiteRoot 'static/downloads/integrations/manifest.json') -Raw | ConvertFrom-Json
$listedPaths = @()
foreach ($artifact in $manifest.artifacts) {
    if (-not $artifact.path.StartsWith('/downloads/integrations/')) {
        throw "Unexpected integration artifact path: $($artifact.path)"
    }
    $path = Resolve-ShowcasePath (Join-Path $WebsiteRoot 'static') $artifact.path.TrimStart('/')
    $listedPaths += $path
    $file = Get-Item -LiteralPath $path -ErrorAction Stop
    if ($file.Length -ne $artifact.bytes -or
        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -cne $artifact.sha256) {
        throw "Integration artifact differs from its manifest: $($artifact.path)"
    }
}
$files = @(Get-ChildItem -LiteralPath (Join-Path $WebsiteRoot 'static/downloads/integrations') -Recurse -File |
    Where-Object Name -NE 'manifest.json' | ForEach-Object FullName)
if (Compare-Object ($files | Sort-Object) ($listedPaths | Sort-Object)) {
    throw 'Integration downloads and manifest entries differ.'
}
foreach ($source in $manifest.articleSources) {
    $path = Resolve-ShowcasePath $WebsiteRoot $source.path
    $text = [IO.File]::ReadAllText($path).Replace("`r`n", "`n")
    $blocks = [regex]::Matches($text, '(?ms)^```powershell\n(.*?)^```')
    $code = ($blocks | ForEach-Object { $_.Groups[1].Value }) -join "`n"
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($code))).ToLowerInvariant()
    if ($hash -cne $source.codeSha256) { throw "Article code differs from the report generation source: $($source.path)" }
}
foreach ($inputFile in $manifest.generationInputs) {
    $path = Resolve-ShowcasePath (Split-Path -Parent $WebsiteRoot) $inputFile.path
    $text = [IO.File]::ReadAllText($path).Replace("`r`n", "`n")
    $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($text))).ToLowerInvariant()
    if ($hash -cne $inputFile.sha256) {
        throw "Report generation input differs from its manifest: $($inputFile.path)"
    }
}
Write-Host "Verified $($manifest.artifacts.Count) integration files and $($manifest.articleSources.Count) article code bindings."
