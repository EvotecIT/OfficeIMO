<# Runs a catalogued PowerShell example and records the actual producer in its child host. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $RepositoryRoot,
    [Parameter(Mandatory)][string] $DocumentsRoot,
    [Parameter(Mandatory)][string] $ExampleId,
    [Parameter(Mandatory)][string] $ModuleVersion,
    [Parameter(Mandatory)][string] $ReceiptPath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ShowcaseEvidence.Helpers.ps1')
. (Join-Path $PSScriptRoot 'ShowcaseEvidence.Provenance.ps1')
$catalog = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'Website/data/showcase.json') -Raw | ConvertFrom-Json -Depth 40
$card = @($catalog.cards | Where-Object id -CEQ $ExampleId)
if ($card.Count -ne 1 -or $card[0].source_language -cne 'powershell') { throw "Not a PowerShell showcase example: $ExampleId" }
$card = $card[0]
$bindings = @(Get-ShowcaseSourceBindings -RepositoryRoot $RepositoryRoot -Catalog $catalog -Card $card)
$producer = Get-ShowcaseEvidenceProducerBinding -RepositoryRoot (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
Import-Module PSWriteOffice -RequiredVersion $ModuleVersion -ErrorAction Stop
$module = Get-Module PSWriteOffice
$packageFiles = @(Get-ChildItem -LiteralPath $module.ModuleBase -File -Recurse | Sort-Object FullName | ForEach-Object {
        [ordered]@{
            path = [IO.Path]::GetRelativePath($module.ModuleBase, $_.FullName).Replace('\', '/')
            sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            fileVersion = $_.VersionInfo.FileVersion
        }
    })
$startedUtc = [datetime]::UtcNow.ToString('o')
$generationRoot = Join-Path $DocumentsRoot ('ShowcaseGeneration-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $generationRoot -ErrorAction Stop | Out-Null
try {
    # The child host owns these variables. Capture the existing exporter's report without changing the recipe.
    $PSDefaultParameterValues['Export-OfficeDocumentPdf:PdfWarningVariable'] = 'global:showcasePdfWarnings'
    $PSDefaultParameterValues['Export-OfficeDocumentPdf:PdfConversionReportVariable'] = 'global:showcasePdfReport'
    & (Resolve-ShowcasePath $RepositoryRoot $card.source_path) -OutputDirectory (Join-Path $generationRoot 'PowerShell') -ModuleVersion $ModuleVersion
    # Hash again so a concurrent edit cannot label output with an input it did not consistently use.
    foreach ($binding in $bindings) {
        if ((Get-FileHash -LiteralPath (Resolve-ShowcasePath $RepositoryRoot $binding.path)).Hash.ToLowerInvariant() -cne $binding.sha256) {
            throw "Showcase source changed during generation: $($binding.path)"
        }
    }
    foreach ($file in $packageFiles) {
        if ((Get-FileHash -LiteralPath (Resolve-ShowcasePath $module.ModuleBase $file.path)).Hash.ToLowerInvariant() -cne $file.sha256) {
            throw "Showcase package changed during generation: $($file.path)"
        }
    }
    foreach ($file in @($producer) + @($producer.dependencies)) {
        if ((Get-FileHash -LiteralPath (Resolve-ShowcasePath (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) $file.path)).Hash.ToLowerInvariant() -cne $file.sha256) {
            throw "Showcase producer changed during generation: $($file.path)"
        }
    }
    $urls = @(Get-ShowcaseCardUrls $card)
    $nativeArtifacts = @($catalog.artifacts | Where-Object {
        $_.sourceRoot -ceq 'documents' -and -not $_.previewSource -and ('/downloads/showcase/' + $_.destination) -cin $urls
    })
    if ($nativeArtifacts.Count -eq 0) { throw "No native outputs declared for '$ExampleId'." }
    $outputs = foreach ($artifact in $nativeArtifacts) {
        $path = Resolve-ShowcasePath $generationRoot $artifact.source
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Current showcase generation did not produce: $($artifact.source)" }
        $diagnostics = $null
        if ([IO.Path]::GetExtension($path) -ceq '.pdf') {
            $info = Get-OfficePdfInfo -Path $path
            $diagnostics = [ordered]@{
                pageCount = $info.PageCount
                fonts = @(Get-OfficePdfFont -Path $path | ForEach-Object {
                    [ordered]@{
                        name = $_.BaseFont
                        subtype = $_.Subtype
                        embedded = $_.HasEmbeddedFontFile
                        programKind = $_.EmbeddedFontFileKind
                        requiresEmbeddingReview = $_.RequiresEmbeddingReview
                        requiresToUnicodeReview = $_.RequiresToUnicodeReview
                        programVersion = $null
                        programSha256 = $null
                        identityStatus = $(if ($_.HasEmbeddedFontFile) { 'not-exposed-by-public-font-diagnostics' } else { 'viewer-resolved-no-embedded-program' })
                    }
                })
            }
        }
        [ordered]@{
            path = '/downloads/showcase/' + $artifact.destination
            sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            diagnostics = $diagnostics
        }
    }
    # Verify the complete set before replacing any existing destination file.
    foreach ($artifact in $nativeArtifacts) {
        $destination = Resolve-ShowcasePath $DocumentsRoot $artifact.source
        New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
        Copy-Item -LiteralPath (Resolve-ShowcasePath $generationRoot $artifact.source) -Destination $destination -Force
    }
    [ordered]@{
        exampleId = $ExampleId
        state = 'generated'
        startedUtc = $startedUtc
        completedUtc = [datetime]::UtcNow.ToString('o')
        host = [ordered]@{ powershell = $PSVersionTable.PSVersion.ToString(); os = [Runtime.InteropServices.RuntimeInformation]::OSDescription; architecture = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() }
        sources = $bindings
        evidenceProducer = $producer
        package = [ordered]@{ name = $module.Name; version = $module.Version.ToString(); fileCount = $packageFiles.Count; files = $packageFiles }
        conversions = @($outputs | Where-Object { $_.path.EndsWith('.pdf', [StringComparison]::Ordinal) } | ForEach-Object {
            [ordered]@{ outputPath = $_.path; outputSha256 = $_.sha256; routeId = 'docx-pdf'; command = 'Export-OfficeDocumentPdf'; routeEvidence = 'recipe-uses-single-document-DOCX-export'; report = $global:showcasePdfReport; warnings = @($global:showcasePdfWarnings) }
        })
        outputs = @($outputs)
    } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $ReceiptPath -Encoding utf8NoBOM
} finally {
    Remove-Item -LiteralPath $generationRoot -Recurse -ErrorAction Stop
}
