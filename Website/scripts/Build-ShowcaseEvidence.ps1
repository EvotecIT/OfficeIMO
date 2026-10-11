<#
.SYNOPSIS
Generates and records downloadable showcase evidence.
.PARAMETER DocumentsRoot
Uses an explicit output directory for PowerShell-only generation, or reads already generated
documents with SkipGeneration. The .NET example runner uses its own Documents directory.
#>
[CmdletBinding()]
param(
    [string] $Framework = 'net10.0',
    [ValidateSet('Debug', 'Release')][string] $Configuration = 'Debug',
    [switch] $SkipGeneration,
    [switch] $ManifestOnly,
    [string[]] $ExampleId = @(),
    [string] $DocumentsRoot = ''
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$customDocumentsRoot = -not [string]::IsNullOrWhiteSpace($DocumentsRoot)
if ([string]::IsNullOrWhiteSpace($DocumentsRoot)) {
    $DocumentsRoot = Join-Path $repoRoot "OfficeIMO.Examples/bin/$Configuration/$Framework/Documents"
}
$downloadRoot = Join-Path $repoRoot 'Website/static/downloads/showcase'
$catalog = Get-Content -LiteralPath (Join-Path $repoRoot 'Website/data/showcase.json') -Raw | ConvertFrom-Json -Depth 40
$moduleVersion = [string](Import-PowerShellDataFile -LiteralPath (Join-Path $repoRoot 'Website/data/apidocs/powershell/PSWriteOffice.psd1')).ModuleVersion
. (Join-Path $PSScriptRoot 'ShowcaseEvidence.Helpers.ps1')

$selectedCards = @($catalog.cards)
if ($ExampleId.Count -gt 0) {
    foreach ($id in $ExampleId) {
        if ($id -cnotin @($catalog.cards.id)) { throw "Unknown showcase example: $id" }
    }
    $selectedCards = @($catalog.cards | Where-Object { $_.id -cin $ExampleId })
}
if ($customDocumentsRoot -and -not $SkipGeneration -and -not $ManifestOnly -and
    @($selectedCards | Where-Object generator_switch).Count -gt 0) {
    throw 'DocumentsRoot supports PowerShell-only generation. For .NET examples, use the default output root or import pre-generated documents with SkipGeneration.'
}
$selectedUrls = @($selectedCards | ForEach-Object {
    @($_.downloads.url) + @($_.source_url, $_.image) + @($_.previews.url)
})
$refreshArtifacts = @($catalog.artifacts | Where-Object {
    $ExampleId.Count -eq 0 -or ('/downloads/showcase/' + $_.destination) -cin $selectedUrls
})

if (-not $SkipGeneration -and -not $ManifestOnly) {
    if (@($selectedCards | Where-Object generator_switch).Count -gt 0) {
        Invoke-ShowcaseDotNet @('build', (Join-Path $repoRoot 'OfficeIMO.Examples/OfficeIMO.Examples.csproj'), '-c', $Configuration, '-f', $Framework, '--nologo')
    }
    $examplesAssembly = Join-Path $repoRoot "OfficeIMO.Examples/bin/$Configuration/$Framework/OfficeIMO.Examples.dll"
    # PowerShell examples are scripts, not switches of the .NET example runner: run each one against the PSWriteOffice
    # version the command catalog was generated from and write its outputs under the shared Documents folder.
    foreach ($card in @($selectedCards | Where-Object { $_.source_path -and $_.source_path.EndsWith('.ps1', [StringComparison]::Ordinal) })) {
        if (-not $moduleVersion) { throw 'The PowerShell command catalog does not name a module version.' }
        $scriptOutput = Join-Path $documentsRoot 'PowerShell'
        & pwsh -NoProfile -File (Resolve-ShowcasePath $repoRoot $card.source_path) -OutputDirectory $scriptOutput -ModuleVersion $moduleVersion
        if ($LASTEXITCODE -ne 0) { throw "PowerShell showcase example failed: $($card.id)" }
    }
    foreach ($exampleSwitch in ($selectedCards.generator_switch | Where-Object { $_ } | Select-Object -Unique)) {
        $generatorCards = @($selectedCards | Where-Object generator_switch -CEQ $exampleSwitch)
        if ($ExampleId.Count -gt 0 -and $exampleSwitch -ceq '--showcase-workflows') {
            foreach ($card in $generatorCards) {
                Invoke-ShowcaseDotNet @($examplesAssembly, $exampleSwitch, '--showcase-example', $card.id)
            }
        } elseif ($ExampleId.Count -gt 0 -and $exampleSwitch -ceq '--showcase-features') {
            foreach ($format in ($generatorCards.format_id | Select-Object -Unique)) {
                Invoke-ShowcaseDotNet @($examplesAssembly, $exampleSwitch, '--showcase-group', $format)
            }
        } else {
            Invoke-ShowcaseDotNet @($examplesAssembly, $exampleSwitch)
        }
    }
}

if (-not $ManifestOnly) {
    $reader = $catalog.artifacts | Where-Object id -eq 'reader-output'
    if ($reader -in $refreshArtifacts) {
        New-ShowcaseReaderProjection `
            -InputPath (Join-Path $documentsRoot 'PowerPoint Design Brief Recommendations.pptx') `
            -OutputPath (Join-Path $documentsRoot $reader.source) `
            -RepositoryRoot $repoRoot -Configuration $Configuration -Framework $Framework
    }
    foreach ($artifact in $refreshArtifacts) {
        if ($artifact.previewSource) {
            New-ShowcasePdfPreview `
                -InputPath (Resolve-ShowcasePath $documentsRoot $artifact.previewSource) `
                -OutputPath (Resolve-ShowcasePath $documentsRoot $artifact.source) `
                -Page ([int] $artifact.previewPage)
        }
    }
}

$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$manifestArtifacts = foreach ($artifact in $catalog.artifacts) {
    if (-not $seen.Add($artifact.destination)) { throw "Duplicate showcase destination: $($artifact.destination)" }
    $destination = Resolve-ShowcasePath $downloadRoot $artifact.destination
    if (-not $ManifestOnly -and $artifact -in $refreshArtifacts) {
        $sourceRoot = switch ($artifact.sourceRoot) {
            'documents' { $documentsRoot }
            'repository' { $repoRoot }
            default { throw "Unknown source root: $($artifact.sourceRoot)" }
        }
        $source = Resolve-ShowcasePath $sourceRoot $artifact.source
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Showcase output missing: $source" }
        New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
        Copy-Item -LiteralPath $source -Destination $destination -Force
    }
    if (-not (Test-Path -LiteralPath $destination -PathType Leaf)) { throw "Showcase download missing: $destination" }
    $file = Get-Item -LiteralPath $destination
    [ordered]@{
        id = $artifact.id
        path = '/downloads/showcase/' + $artifact.destination
        bytes = $file.Length
        sha256 = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant()
        generator = $artifact.generator.Replace('-f net10.0', "-f $Framework").Replace('{moduleVersion}', $moduleVersion)
        evidence = $artifact.evidence.Replace('{moduleVersion}', $moduleVersion)
    }
}

foreach ($card in $catalog.cards) {
    foreach ($url in @($card.downloads.url) + @($card.source_url, $card.image) + @($card.previews.url)) {
        if ($url -and (-not $url.StartsWith('/downloads/showcase/') -or
            -not $seen.Contains($url.Substring('/downloads/showcase/'.Length)))) {
            throw "Example '$($card.id)' references an undeclared artifact: $url"
        }
    }
}
[ordered]@{
    schema = 'officeimo.showcase-evidence'
    schemaVersion = 1
    artifacts = @($manifestArtifacts)
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $downloadRoot 'manifest.json') -Encoding utf8NoBOM
Write-Host "Showcase evidence refreshed: $($catalog.cards.Count) examples, $($manifestArtifacts.Count) artifacts."
& (Join-Path $PSScriptRoot 'Sync-ShowcaseWalkthroughs.ps1')
