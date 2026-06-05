[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $ShowcasePath,

    [switch] $RequirePreviewsPerDiagram,

    [switch] $RequireProofsPerDiagram,

    [switch] $RequireNativePreviewFormatsPerDiagram,

    [switch] $RequireDesktopPreviewFormatsPerDiagram
)

$ErrorActionPreference = 'Stop'

function Assert-Condition {
    param(
        [Parameter(Mandatory = $true)]
        [bool] $Condition,

        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

. (Join-Path $PSScriptRoot 'Assert-VisioShowcaseCommon.ps1')
. (Join-Path $PSScriptRoot 'Assert-VisioShowcaseVisualQualityProof.ps1')
. (Join-Path $PSScriptRoot 'Assert-VisioShowcaseEvidenceTotals.ps1')

function Assert-GalleryContains {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Needle,

        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    Assert-Condition -Condition ($script:GalleryHtml.Contains($Needle)) -Message $Message
}

function Assert-MarkdownContains {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Needle,

        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    Assert-Condition -Condition ($script:SummaryMarkdown.Contains($Needle)) -Message $Message
}

function Assert-GalleryArtifactLink {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Artifact
    )

    $href = ConvertTo-GalleryHref -RelativePath $Artifact.relativePath
    Assert-GalleryContains -Needle "href=""$href""" -Message "Gallery is missing a link for artifact: $($Artifact.relativePath)"
    Assert-GalleryContains -Needle "title=""$($Artifact.sha256)""" -Message "Gallery is missing SHA-256 proof for artifact: $($Artifact.relativePath)"
}

function Assert-DiagramProofSummary {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Diagram,

        [Parameter(Mandatory = $true)]
        [object[]] $DiagramProofs
    )

    Assert-Condition -Condition ($Diagram.PSObject.Properties.Name -contains 'proofSummary') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary."
    $proofSummary = $Diagram.proofSummary
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'totalShapeCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.totalShapeCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'connectorCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.connectorCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'stencilCatalogCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.stencilCatalogCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'stencilUsageCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.stencilUsageCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'shapeDataKeyCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.shapeDataKeyCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'connectorShapeDataKeyCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.connectorShapeDataKeyCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'semanticKindCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.semanticKindCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'totalConnectionPointCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.totalConnectionPointCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'connectionPointShapeCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.connectionPointShapeCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'stencilFamilyCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.stencilFamilyCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'stencilBackedShapeCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.stencilBackedShapeCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'basicGeometryShapeCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.basicGeometryShapeCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'masterBackedShapeCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.masterBackedShapeCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'packageBackedShapeCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.packageBackedShapeCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'generatedMasterBackedShapeCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.generatedMasterBackedShapeCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'semanticOnlyShapeCount') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.semanticOnlyShapeCount."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'stencilCatalogs') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.stencilCatalogs."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'shapeDataKeys') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.shapeDataKeys."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'connectorShapeDataKeys') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.connectorShapeDataKeys."
    Assert-Condition -Condition ($proofSummary.PSObject.Properties.Name -contains 'semanticKinds') -Message "Showcase diagram '$($Diagram.name)' is missing proofSummary.semanticKinds."

    $stencilCatalogs = Get-RequiredArray -Value $proofSummary.stencilCatalogs
    $shapeDataKeys = Get-RequiredArray -Value $proofSummary.shapeDataKeys
    $connectorShapeDataKeys = Get-RequiredArray -Value $proofSummary.connectorShapeDataKeys
    $semanticKinds = Get-RequiredArray -Value $proofSummary.semanticKinds
    Assert-Condition -Condition ($proofSummary.totalShapeCount -is [int] -or $proofSummary.totalShapeCount -is [long]) -Message "Showcase diagram '$($Diagram.name)' totalShapeCount is not numeric."
    Assert-Condition -Condition ([long] $proofSummary.totalShapeCount -ge 0) -Message "Showcase diagram '$($Diagram.name)' totalShapeCount cannot be negative."
    Assert-Condition -Condition ($proofSummary.connectorCount -is [int] -or $proofSummary.connectorCount -is [long]) -Message "Showcase diagram '$($Diagram.name)' connectorCount is not numeric."
    Assert-Condition -Condition ([long] $proofSummary.connectorCount -ge 0) -Message "Showcase diagram '$($Diagram.name)' connectorCount cannot be negative."
    if ([long] $proofSummary.totalShapeCount -gt 0) {
        Assert-Condition -Condition ([long] $proofSummary.connectorCount -le [long] $proofSummary.totalShapeCount) -Message "Showcase diagram '$($Diagram.name)' connectorCount cannot exceed totalShapeCount."
    }

    Assert-Condition -Condition ($proofSummary.stencilCatalogCount -eq $stencilCatalogs.Count) -Message "Showcase diagram '$($Diagram.name)' stencilCatalogCount mismatch. summary=$($proofSummary.stencilCatalogCount), catalogs=$($stencilCatalogs.Count)."
    Assert-Condition -Condition ($proofSummary.stencilUsageCount -is [int] -or $proofSummary.stencilUsageCount -is [long]) -Message "Showcase diagram '$($Diagram.name)' stencilUsageCount is not numeric."
    Assert-Condition -Condition ([long] $proofSummary.stencilUsageCount -ge 0) -Message "Showcase diagram '$($Diagram.name)' stencilUsageCount cannot be negative."
    Assert-Condition -Condition ($proofSummary.shapeDataKeyCount -eq $shapeDataKeys.Count) -Message "Showcase diagram '$($Diagram.name)' shapeDataKeyCount mismatch. summary=$($proofSummary.shapeDataKeyCount), keys=$($shapeDataKeys.Count)."
    Assert-Condition -Condition ($proofSummary.connectorShapeDataKeyCount -eq $connectorShapeDataKeys.Count) -Message "Showcase diagram '$($Diagram.name)' connectorShapeDataKeyCount mismatch. summary=$($proofSummary.connectorShapeDataKeyCount), keys=$($connectorShapeDataKeys.Count)."
    Assert-Condition -Condition ($proofSummary.semanticKindCount -eq $semanticKinds.Count) -Message "Showcase diagram '$($Diagram.name)' semanticKindCount mismatch. summary=$($proofSummary.semanticKindCount), semanticKinds=$($semanticKinds.Count)."
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.totalConnectionPointCount -Label 'totalConnectionPointCount'
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.connectionPointShapeCount -Label 'connectionPointShapeCount'
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.stencilFamilyCount -Label 'stencilFamilyCount'
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.stencilBackedShapeCount -Label 'stencilBackedShapeCount'
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.basicGeometryShapeCount -Label 'basicGeometryShapeCount'
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.masterBackedShapeCount -Label 'masterBackedShapeCount'
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.packageBackedShapeCount -Label 'packageBackedShapeCount'
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.generatedMasterBackedShapeCount -Label 'generatedMasterBackedShapeCount'
    Assert-ProofSummaryNonNegativeNumber -Diagram $Diagram -Value $proofSummary.semanticOnlyShapeCount -Label 'semanticOnlyShapeCount'
    if ([long] $proofSummary.totalShapeCount -gt 0) {
        Assert-Condition -Condition ([long] $proofSummary.connectionPointShapeCount -le [long] $proofSummary.totalShapeCount) -Message "Showcase diagram '$($Diagram.name)' connectionPointShapeCount cannot exceed totalShapeCount."
        Assert-Condition -Condition ([long] $proofSummary.stencilBackedShapeCount -le [long] $proofSummary.totalShapeCount) -Message "Showcase diagram '$($Diagram.name)' stencilBackedShapeCount cannot exceed totalShapeCount."
    }

    Assert-GalleryContains -Needle "$($proofSummary.totalShapeCount) shapes / $($proofSummary.connectorCount) connectors / $($proofSummary.shapeDataKeyCount) shape-data keys" -Message "Gallery is missing structural proof metrics for diagram '$($Diagram.name)'."
    Assert-GalleryContains -Needle "$($proofSummary.stencilBackedShapeCount) stencil-backed / $($proofSummary.basicGeometryShapeCount) basic geometry / $($proofSummary.totalConnectionPointCount) connection points" -Message "Gallery is missing stencil backing metrics for diagram '$($Diagram.name)'."

    $catalogLookup = @{}
    foreach ($catalog in $stencilCatalogs) {
        Assert-Condition -Condition ($catalog -is [string] -and -not [string]::IsNullOrWhiteSpace($catalog)) -Message "Showcase diagram '$($Diagram.name)' has an empty stencil catalog."
        $catalogKey = $catalog.ToLowerInvariant()
        Assert-Condition -Condition (-not $catalogLookup.ContainsKey($catalogKey)) -Message "Showcase diagram '$($Diagram.name)' has duplicate stencil catalog '$catalog'."
        $catalogLookup[$catalogKey] = $true
        Assert-GalleryContains -Needle ([System.Net.WebUtility]::HtmlEncode($catalog)) -Message "Gallery is missing stencil catalog '$catalog' for diagram '$($Diagram.name)'."
        Assert-MarkdownContains -Needle $catalog -Message "Markdown summary is missing stencil catalog '$catalog' for diagram '$($Diagram.name)'."
    }

    Assert-ProofSummaryStringArray -Diagram $Diagram -Values $shapeDataKeys -Label 'shape-data key'
    Assert-ProofSummaryStringArray -Diagram $Diagram -Values $connectorShapeDataKeys -Label 'connector shape-data key'
    Assert-ProofSummaryStringArray -Diagram $Diagram -Values $semanticKinds -Label 'semantic kind'

    $hasStencilProfile = @($DiagramProofs | Where-Object { $_.kind -eq 'StencilProfile' }).Count -gt 0
    if ($hasStencilProfile) {
        Assert-MarkdownContains -Needle "- $($Diagram.name):" -Message "Markdown summary is missing proof summary line for diagram '$($Diagram.name)'."
        Assert-MarkdownContains -Needle "$($proofSummary.totalShapeCount) shapes" -Message "Markdown summary is missing structural shape count for diagram '$($Diagram.name)'."
        Assert-MarkdownContains -Needle "$($proofSummary.connectorCount) connectors" -Message "Markdown summary is missing structural connector count for diagram '$($Diagram.name)'."
        Assert-MarkdownContains -Needle "$($proofSummary.stencilBackedShapeCount) stencil-backed shapes" -Message "Markdown summary is missing stencil-backed shape count for diagram '$($Diagram.name)'."
        Assert-MarkdownContains -Needle "$($proofSummary.totalConnectionPointCount) connection points" -Message "Markdown summary is missing connection point count for diagram '$($Diagram.name)'."
    }
}

function Assert-ProofSummaryNonNegativeNumber {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Diagram,

        [Parameter(Mandatory = $true)]
        [object] $Value,

        [Parameter(Mandatory = $true)]
        [string] $Label
    )

    Assert-Condition -Condition ($Value -is [int] -or $Value -is [long]) -Message "Showcase diagram '$($Diagram.name)' $Label is not numeric."
    Assert-Condition -Condition ([long] $Value -ge 0) -Message "Showcase diagram '$($Diagram.name)' $Label cannot be negative."
}

function Assert-ProofSummaryStringArray {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Diagram,

        [object[]] $Values = @(),

        [Parameter(Mandatory = $true)]
        [string] $Label
    )

    $lookup = @{}
    foreach ($value in $Values) {
        Assert-Condition -Condition ($value -is [string] -and -not [string]::IsNullOrWhiteSpace($value)) -Message "Showcase diagram '$($Diagram.name)' has an empty $Label."
        $key = $value.ToLowerInvariant()
        Assert-Condition -Condition (-not $lookup.ContainsKey($key)) -Message "Showcase diagram '$($Diagram.name)' has duplicate $Label '$value'."
        $lookup[$key] = $true
    }
}

function Assert-ProofTotals {
    Assert-Condition -Condition ($summary.PSObject.Properties.Name -contains 'proofTotals') -Message 'Summary is missing proofTotals.'
    $totals = $summary.proofTotals
    foreach ($propertyName in @(
        'totalShapeCount',
        'connectorCount',
        'stencilUsageCount',
        'totalConnectionPointCount',
        'connectionPointShapeCount',
        'stencilFamilyCount',
        'stencilBackedShapeCount',
        'basicGeometryShapeCount',
        'masterBackedShapeCount',
        'packageBackedShapeCount',
        'generatedMasterBackedShapeCount',
        'semanticOnlyShapeCount',
        'stencilCatalogCount',
        'shapeDataKeyCount',
        'connectorShapeDataKeyCount',
        'semanticKindCount')) {
        Assert-Condition -Condition ($totals.PSObject.Properties.Name -contains $propertyName) -Message "proofTotals is missing '$propertyName'."
        Assert-Condition -Condition ($totals.$propertyName -is [int] -or $totals.$propertyName -is [long]) -Message "proofTotals.$propertyName is not numeric."
        Assert-Condition -Condition ([long] $totals.$propertyName -ge 0) -Message "proofTotals.$propertyName cannot be negative."
    }

    $stencilCatalogs = Get-RequiredArray -Value $totals.stencilCatalogs
    $shapeDataKeys = Get-RequiredArray -Value $totals.shapeDataKeys
    $connectorShapeDataKeys = Get-RequiredArray -Value $totals.connectorShapeDataKeys
    $semanticKinds = Get-RequiredArray -Value $totals.semanticKinds
    Assert-Condition -Condition ($totals.stencilCatalogCount -eq $stencilCatalogs.Count) -Message "proofTotals.stencilCatalogCount mismatch. summary=$($totals.stencilCatalogCount), values=$($stencilCatalogs.Count)."
    Assert-Condition -Condition ($totals.shapeDataKeyCount -eq $shapeDataKeys.Count) -Message "proofTotals.shapeDataKeyCount mismatch. summary=$($totals.shapeDataKeyCount), values=$($shapeDataKeys.Count)."
    Assert-Condition -Condition ($totals.connectorShapeDataKeyCount -eq $connectorShapeDataKeys.Count) -Message "proofTotals.connectorShapeDataKeyCount mismatch. summary=$($totals.connectorShapeDataKeyCount), values=$($connectorShapeDataKeys.Count)."
    Assert-Condition -Condition ($totals.semanticKindCount -eq $semanticKinds.Count) -Message "proofTotals.semanticKindCount mismatch. summary=$($totals.semanticKindCount), values=$($semanticKinds.Count)."

    $expected = [ordered]@{
        totalShapeCount = 0L
        connectorCount = 0L
        stencilUsageCount = 0L
        totalConnectionPointCount = 0L
        connectionPointShapeCount = 0L
        stencilFamilyCount = 0L
        stencilBackedShapeCount = 0L
        basicGeometryShapeCount = 0L
        masterBackedShapeCount = 0L
        packageBackedShapeCount = 0L
        generatedMasterBackedShapeCount = 0L
        semanticOnlyShapeCount = 0L
    }
    $catalogLookup = @{}
    $shapeDataLookup = @{}
    $connectorShapeDataLookup = @{}
    $semanticKindLookup = @{}
    foreach ($diagram in $diagrams) {
        $proofSummary = $diagram.proofSummary
        foreach ($propertyName in @($expected.Keys)) {
            $expected[$propertyName] += [long] $proofSummary.$propertyName
        }

        foreach ($catalog in (Get-RequiredArray -Value $proofSummary.stencilCatalogs)) {
            $catalogLookup[$catalog.ToLowerInvariant()] = $catalog
        }

        foreach ($key in (Get-RequiredArray -Value $proofSummary.shapeDataKeys)) {
            $shapeDataLookup[$key.ToLowerInvariant()] = $key
        }

        foreach ($key in (Get-RequiredArray -Value $proofSummary.connectorShapeDataKeys)) {
            $connectorShapeDataLookup[$key.ToLowerInvariant()] = $key
        }

        foreach ($kind in (Get-RequiredArray -Value $proofSummary.semanticKinds)) {
            $semanticKindLookup[$kind.ToLowerInvariant()] = $kind
        }
    }

    foreach ($propertyName in @($expected.Keys)) {
        Assert-Condition -Condition ([long] $totals.$propertyName -eq $expected[$propertyName]) -Message "proofTotals.$propertyName mismatch. summary=$($totals.$propertyName), diagrams=$($expected[$propertyName])."
    }

    Assert-Condition -Condition ($totals.stencilCatalogCount -eq $catalogLookup.Count) -Message "proofTotals.stencilCatalogCount does not match distinct diagram catalogs."
    Assert-Condition -Condition ($totals.shapeDataKeyCount -eq $shapeDataLookup.Count) -Message "proofTotals.shapeDataKeyCount does not match distinct diagram Shape Data keys."
    Assert-Condition -Condition ($totals.connectorShapeDataKeyCount -eq $connectorShapeDataLookup.Count) -Message "proofTotals.connectorShapeDataKeyCount does not match distinct diagram connector Shape Data keys."
    Assert-Condition -Condition ($totals.semanticKindCount -eq $semanticKindLookup.Count) -Message "proofTotals.semanticKindCount does not match distinct diagram semantic kinds."
    Assert-MarkdownContains -Needle "Proof total shapes: $($totals.totalShapeCount)" -Message 'Markdown summary proof total shapes are missing or stale.'
    Assert-MarkdownContains -Needle "Proof total connectors: $($totals.connectorCount)" -Message 'Markdown summary proof total connectors are missing or stale.'
    Assert-MarkdownContains -Needle "Proof stencil-backed shapes: $($totals.stencilBackedShapeCount)" -Message 'Markdown summary proof stencil-backed shapes are missing or stale.'
    Assert-MarkdownContains -Needle "Proof connection points: $($totals.totalConnectionPointCount)" -Message 'Markdown summary proof connection points are missing or stale.'
    Assert-MarkdownContains -Needle '## Proof Totals' -Message 'Markdown summary is missing the proof totals section.'
    Assert-MarkdownContains -Needle "$($totals.totalShapeCount) shapes, $($totals.connectorCount) connectors, $($totals.stencilBackedShapeCount) stencil-backed shapes" -Message 'Markdown summary proof totals line is missing or stale.'
    Assert-GalleryContains -Needle 'href="#proof-totals"' -Message 'Gallery proof totals navigation link is missing.'
    Assert-GalleryContains -Needle '<h2 id="proof-totals">Proof Totals</h2>' -Message 'Gallery proof totals section is missing.'
    Assert-GalleryContains -Needle "<tr><td>Shapes</td><td>$($totals.totalShapeCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td></tr>" -Message 'Gallery proof total shape row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Connectors</td><td>$($totals.connectorCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td></tr>" -Message 'Gallery proof total connector row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Stencil-backed shapes</td><td>$($totals.stencilBackedShapeCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td></tr>" -Message 'Gallery proof total stencil-backed shape row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Connection points</td><td>$($totals.totalConnectionPointCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td></tr>" -Message 'Gallery proof total connection-point row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Distinct Shape Data keys</td><td>$($totals.shapeDataKeyCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td></tr>" -Message 'Gallery proof total Shape Data row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Distinct connector Shape Data keys</td><td>$($totals.connectorShapeDataKeyCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td></tr>" -Message 'Gallery proof total connector Shape Data row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Semantic kinds</td><td>$($totals.semanticKindCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td></tr>" -Message 'Gallery proof total semantic-kind row is missing or stale.'
    Assert-GalleryContains -Needle "<strong>$($totals.totalShapeCount)</strong><span>Proof shapes</span>" -Message 'Gallery proof shapes metric is missing or stale.'
    Assert-GalleryContains -Needle "<strong>$($totals.connectorCount)</strong><span>Proof connectors</span>" -Message 'Gallery proof connectors metric is missing or stale.'
    Assert-GalleryContains -Needle "<strong>$($totals.stencilBackedShapeCount)</strong><span>Stencil-backed shapes</span>" -Message 'Gallery stencil-backed shapes metric is missing or stale.'
    Assert-GalleryContains -Needle "<strong>$($totals.totalConnectionPointCount)</strong><span>Connection points</span>" -Message 'Gallery connection points metric is missing or stale.'
}

function Test-DiagramPreviewEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Diagram,

        [Parameter(Mandatory = $true)]
        [string] $Kind,

        [Parameter(Mandatory = $true)]
        [string] $Format
    )

    foreach ($preview in (Get-RequiredArray -Value $Diagram.previews)) {
        if ($preview.kind -eq $Kind -and $preview.format -eq $Format) {
            return $true
        }
    }

    return $false
}

function Test-DiagramProofEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Diagram,

        [Parameter(Mandatory = $true)]
        [string] $Kind
    )

    foreach ($proof in (Get-RequiredArray -Value $Diagram.proofs)) {
        if ($proof.kind -eq $Kind) {
            return $true
        }
    }

    return $false
}

function Assert-Artifact {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Artifact
    )

    Assert-Condition -Condition ($Artifact.kind -is [string]) -Message 'Artifact kind is missing.'
    Assert-Condition -Condition ($Artifact.format -is [string]) -Message "Artifact '$($Artifact.relativePath)' format is missing."
    Assert-Condition -Condition ($Artifact.relativePath -is [string]) -Message 'Artifact relativePath is missing.'
    Assert-Condition -Condition ($Artifact.sizeBytes -is [int] -or $Artifact.sizeBytes -is [long]) -Message "Artifact '$($Artifact.relativePath)' sizeBytes is missing."
    Assert-Condition -Condition ($Artifact.sha256 -is [string] -and $Artifact.sha256 -match '^[a-f0-9]{64}$') -Message "Artifact '$($Artifact.relativePath)' is missing a valid SHA-256 hash."

    $artifactPath = Resolve-ShowcaseArtifactPath -RelativePath $Artifact.relativePath
    Assert-Condition -Condition (Test-Path -LiteralPath $artifactPath -PathType Leaf) -Message "Artifact file is missing: $($Artifact.relativePath)"

    $file = Get-Item -LiteralPath $artifactPath
    Assert-Condition -Condition ($file.Length -eq [long] $Artifact.sizeBytes) -Message "Artifact '$($Artifact.relativePath)' size mismatch. summary=$($Artifact.sizeBytes), file=$($file.Length)."

    $actualHash = (Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-Condition -Condition ($actualHash -eq $Artifact.sha256) -Message "Artifact '$($Artifact.relativePath)' SHA-256 mismatch. summary=$($Artifact.sha256), file=$actualHash."
}

function Assert-GallerySurface {
    Assert-GalleryContains -Needle '<title>OfficeIMO Visio Showcase</title>' -Message 'Gallery HTML title is missing.'
    Assert-GalleryContains -Needle 'href="showcase-summary.json"' -Message 'Gallery is missing the JSON summary link.'
    Assert-GalleryContains -Needle 'href="showcase-summary.md"' -Message 'Gallery is missing the Markdown summary link.'

    $sections = @(
        @{ Id = 'review-index'; Label = 'Review Index' },
        @{ Id = 'stencil-coverage'; Label = 'Stencil Coverage' },
        @{ Id = 'diagram-review-cards'; Label = 'Diagram Review Cards' },
        @{ Id = 'packages'; Label = 'Packages' },
        @{ Id = 'previews'; Label = 'Previews' },
        @{ Id = 'structural-proofs'; Label = 'Structural Proofs' }
    )

    foreach ($section in $sections) {
        Assert-GalleryContains -Needle "href=""#$($section.Id)""" -Message "Gallery navigation is missing the '$($section.Label)' link."
        Assert-GalleryContains -Needle "<h2 id=""$($section.Id)"">$($section.Label)</h2>" -Message "Gallery is missing the '$($section.Label)' section."
    }

    $metrics = @(
        @{ Label = 'VSDX packages'; Value = $summary.packageCount },
        @{ Label = 'Preview files'; Value = $summary.previewCount },
        @{ Label = 'Structural proofs'; Value = $summary.proofCount },
        @{ Label = 'Stencil catalogs'; Value = $summary.stencilCatalogCount },
        @{ Label = 'Total artifacts'; Value = $summary.artifactCount }
    )

    foreach ($metric in $metrics) {
        Assert-GalleryContains -Needle "<strong>$($metric.Value)</strong><span>$($metric.Label)</span>" -Message "Gallery metric '$($metric.Label)' is missing or stale."
    }

    $hashTitleCount = [regex]::Matches($script:GalleryHtml, 'title="[a-f0-9]{64}"').Count
    Assert-Condition -Condition ($hashTitleCount -ge $artifacts.Count) -Message "Gallery exposes too few full SHA-256 hash titles. expected-at-least=$($artifacts.Count), actual=$hashTitleCount."

    foreach ($artifact in $artifacts) {
        Assert-GalleryArtifactLink -Artifact $artifact
    }

    foreach ($diagram in $diagrams) {
        $diagramFragmentId = 'diagram-' + (ConvertTo-GalleryFragmentId -Value $diagram.package.relativePath)
        Assert-GalleryContains -Needle "id=""$diagramFragmentId""" -Message "Gallery is missing the review card anchor for diagram: $($diagram.name)"
        Assert-GalleryContains -Needle "href=""#$diagramFragmentId""" -Message "Gallery review index is missing the deep link for diagram: $($diagram.name)"
    }

    Assert-GalleryContains -Needle '<th>Stencil Catalogs</th>' -Message 'Gallery review index is missing the stencil-catalog summary column.'
    Assert-GalleryContains -Needle '<th>Stencil Catalog</th><th>Diagrams</th><th>Diagram Names</th>' -Message 'Gallery stencil coverage table is missing or stale.'
}

function Assert-MarkdownSurface {
    Assert-MarkdownContains -Needle '# OfficeIMO Visio Showcase Summary' -Message 'Markdown summary title is missing.'
    Assert-MarkdownContains -Needle "Diagrams: $($summary.diagramCount)" -Message 'Markdown summary diagram count is missing or stale.'
    Assert-MarkdownContains -Needle "VSDX files: $($summary.packageCount)" -Message 'Markdown summary package count is missing or stale.'
    Assert-MarkdownContains -Needle "Preview files: $($summary.previewCount)" -Message 'Markdown summary preview count is missing or stale.'
    Assert-MarkdownContains -Needle "Structural proof files: $($summary.proofCount)" -Message 'Markdown summary proof count is missing or stale.'
    Assert-MarkdownContains -Needle "Stencil catalogs: $($summary.stencilCatalogCount)" -Message 'Markdown summary stencil catalog count is missing or stale.'
    Assert-MarkdownContains -Needle "Total artifacts: $($summary.artifactCount)" -Message 'Markdown summary artifact count is missing or stale.'
    Assert-MarkdownContains -Needle 'Machine-readable summary: `showcase-summary.json`' -Message 'Markdown summary is missing the JSON summary pointer.'
    Assert-MarkdownContains -Needle 'Browsable gallery: `showcase-gallery.html`' -Message 'Markdown summary is missing the HTML gallery pointer.'
    Assert-MarkdownContains -Needle '## Stencil Catalog Coverage' -Message 'Markdown summary is missing the stencil catalog coverage section.'
    Assert-MarkdownContains -Needle '## Diagram Proof Summary' -Message 'Markdown summary is missing the diagram proof summary section.'

    foreach ($artifact in $artifacts) {
        Assert-MarkdownContains -Needle "``$($artifact.relativePath)``" -Message "Markdown summary is missing artifact path: $($artifact.relativePath)"
        Assert-MarkdownContains -Needle "sha256: $($artifact.sha256)" -Message "Markdown summary is missing SHA-256 proof for artifact: $($artifact.relativePath)"
    }
}

$script:ShowcaseRoot = [System.IO.Path]::GetFullPath($ShowcasePath)
$script:ShowcaseRootWithSeparator = $script:ShowcaseRoot
if (-not $script:ShowcaseRootWithSeparator.EndsWith([System.IO.Path]::DirectorySeparatorChar.ToString(), [StringComparison]::Ordinal)) {
    $script:ShowcaseRootWithSeparator += [System.IO.Path]::DirectorySeparatorChar
}

Assert-Condition -Condition (Test-Path -LiteralPath $script:ShowcaseRoot -PathType Container) -Message "Showcase folder is missing: $script:ShowcaseRoot"

$summaryPath = Join-Path $script:ShowcaseRoot 'showcase-summary.json'
$galleryPath = Join-Path $script:ShowcaseRoot 'showcase-gallery.html'
$markdownPath = Join-Path $script:ShowcaseRoot 'showcase-summary.md'

foreach ($path in @($summaryPath, $galleryPath, $markdownPath)) {
    Assert-Condition -Condition (Test-Path -LiteralPath $path -PathType Leaf) -Message "Required showcase artifact is missing: $path"
}

$summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
$script:GalleryHtml = Get-Content -LiteralPath $galleryPath -Raw
$script:SummaryMarkdown = Get-Content -LiteralPath $markdownPath -Raw
$diagrams = Get-RequiredArray -Value $summary.diagrams
$artifacts = Get-RequiredArray -Value $summary.artifacts
$packages = @($artifacts | Where-Object { $_.kind -eq 'Package' })
$previews = @($artifacts | Where-Object { $_.kind -in @('NativePreview', 'DesktopPreview', 'Preview') })
$proofs = @($artifacts | Where-Object { $_.kind -in @('Inspection', 'StencilProfile', 'VisualQuality', 'Proof') })
$summaryStencilCatalogs = Get-RequiredArray -Value $summary.stencilCatalogs
$summaryStencilCatalogCoverage = Get-RequiredArray -Value $summary.stencilCatalogCoverage

Assert-Condition -Condition ($summary.schemaVersion -eq 1) -Message "Unsupported showcase summary schemaVersion '$($summary.schemaVersion)'."
Assert-Condition -Condition ($summary.diagramCount -eq $diagrams.Count) -Message "diagramCount $($summary.diagramCount) does not match diagrams array count $($diagrams.Count)."
Assert-Condition -Condition ($summary.artifactCount -eq $artifacts.Count) -Message "artifactCount $($summary.artifactCount) does not match artifacts array count $($artifacts.Count)."
Assert-Condition -Condition ($summary.packageCount -eq $packages.Count -and $summary.packageCount -eq $diagrams.Count) -Message "Package count mismatch. summary=$($summary.packageCount), artifacts=$($packages.Count), diagrams=$($diagrams.Count)."
Assert-Condition -Condition ($summary.previewCount -eq $previews.Count) -Message "Preview count mismatch. summary=$($summary.previewCount), artifacts=$($previews.Count)."
Assert-Condition -Condition ($summary.proofCount -eq $proofs.Count) -Message "Proof count mismatch. summary=$($summary.proofCount), artifacts=$($proofs.Count)."
Assert-Condition -Condition ($summary.stencilCatalogCount -eq $summaryStencilCatalogs.Count) -Message "stencilCatalogCount $($summary.stencilCatalogCount) does not match stencilCatalogs array count $($summaryStencilCatalogs.Count)."
Assert-Condition -Condition ($summaryStencilCatalogCoverage.Count -eq $summaryStencilCatalogs.Count) -Message "stencilCatalogCoverage count $($summaryStencilCatalogCoverage.Count) does not match stencilCatalogs array count $($summaryStencilCatalogs.Count)."
Assert-ProofTotals
Assert-EvidenceTotals

$summaryCatalogLookup = @{}
foreach ($catalog in $summaryStencilCatalogs) {
    Assert-Condition -Condition ($catalog -is [string] -and -not [string]::IsNullOrWhiteSpace($catalog)) -Message 'Summary stencilCatalogs contains an empty value.'
    $catalogKey = $catalog.ToLowerInvariant()
    Assert-Condition -Condition (-not $summaryCatalogLookup.ContainsKey($catalogKey)) -Message "Summary stencilCatalogs contains duplicate value '$catalog'."
    $summaryCatalogLookup[$catalogKey] = $true
}

$coverageCatalogLookup = @{}
foreach ($coverage in $summaryStencilCatalogCoverage) {
    Assert-Condition -Condition ($coverage.catalog -is [string] -and -not [string]::IsNullOrWhiteSpace($coverage.catalog)) -Message 'stencilCatalogCoverage contains an empty catalog.'
    $coverageCatalogKey = $coverage.catalog.ToLowerInvariant()
    Assert-Condition -Condition ($summaryCatalogLookup.ContainsKey($coverageCatalogKey)) -Message "stencilCatalogCoverage has catalog '$($coverage.catalog)' that is missing from stencilCatalogs."
    Assert-Condition -Condition (-not $coverageCatalogLookup.ContainsKey($coverageCatalogKey)) -Message "stencilCatalogCoverage contains duplicate catalog '$($coverage.catalog)'."
    $coverageCatalogLookup[$coverageCatalogKey] = $coverage

    $coverageDiagrams = Get-RequiredArray -Value $coverage.diagrams
    Assert-Condition -Condition ($coverage.diagramCount -eq $coverageDiagrams.Count) -Message "stencilCatalogCoverage diagramCount mismatch for '$($coverage.catalog)'. summary=$($coverage.diagramCount), diagrams=$($coverageDiagrams.Count)."
    foreach ($diagramName in $coverageDiagrams) {
        Assert-Condition -Condition ($diagramName -is [string] -and -not [string]::IsNullOrWhiteSpace($diagramName)) -Message "stencilCatalogCoverage '$($coverage.catalog)' contains an empty diagram name."
        Assert-MarkdownContains -Needle $diagramName -Message "Markdown summary is missing stencil coverage diagram '$diagramName'."
        Assert-GalleryContains -Needle ([System.Net.WebUtility]::HtmlEncode($diagramName)) -Message "Gallery is missing stencil coverage diagram '$diagramName'."
    }
}

$artifactsByPath = @{}
foreach ($artifact in $artifacts) {
    Assert-Artifact -Artifact $artifact
    Assert-Condition -Condition (-not $artifactsByPath.ContainsKey($artifact.relativePath)) -Message "Duplicate artifact relativePath: $($artifact.relativePath)"
    $artifactsByPath[$artifact.relativePath] = $artifact
}

$diagramPackagePaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$diagramPreviewPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$diagramProofPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

foreach ($diagram in $diagrams) {
    $diagramPreviews = Get-RequiredArray -Value $diagram.previews
    $diagramProofs = Get-RequiredArray -Value $diagram.proofs
    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($diagram.name)) -Message 'A showcase diagram is missing its name.'
    Assert-Condition -Condition ($diagram.package.relativePath -is [string]) -Message "Showcase diagram '$($diagram.name)' is missing its package path."
    Assert-Condition -Condition ($diagram.previewCount -eq $diagramPreviews.Count) -Message "Preview count mismatch for '$($diagram.name)'. summary=$($diagram.previewCount), previews=$($diagramPreviews.Count)."
    Assert-Condition -Condition ($diagram.proofCount -eq $diagramProofs.Count) -Message "Proof count mismatch for '$($diagram.name)'. summary=$($diagram.proofCount), proofs=$($diagramProofs.Count)."
    Assert-DiagramProofSummary -Diagram $diagram -DiagramProofs $diagramProofs
    if ($RequirePreviewsPerDiagram) {
        Assert-Condition -Condition ($diagramPreviews.Count -gt 0) -Message "Showcase diagram '$($diagram.name)' has no preview artifacts."
    }

    if ($RequireProofsPerDiagram) {
        Assert-Condition -Condition ($diagramProofs.Count -gt 0) -Message "Showcase diagram '$($diagram.name)' has no structural proof artifacts."
    }

    $packagePath = $diagram.package.relativePath
    Assert-Condition -Condition ($artifactsByPath.ContainsKey($packagePath)) -Message "Showcase diagram '$($diagram.name)' package '$packagePath' is not listed in artifacts."
    $packageArtifact = $artifactsByPath[$packagePath]
    Assert-Condition -Condition ($packageArtifact.kind -eq 'Package') -Message "Showcase diagram '$($diagram.name)' package '$packagePath' is not a package artifact."
    Assert-Condition -Condition ($diagram.package.sha256 -eq $packageArtifact.sha256) -Message "Showcase diagram '$($diagram.name)' package hash does not match the artifact list."
    Assert-Condition -Condition ($diagram.package.sizeBytes -eq $packageArtifact.sizeBytes) -Message "Showcase diagram '$($diagram.name)' package size does not match the artifact list."
    [void] $diagramPackagePaths.Add($packagePath)

    foreach ($preview in $diagramPreviews) {
        $previewPath = $preview.relativePath
        Assert-Condition -Condition ($artifactsByPath.ContainsKey($previewPath)) -Message "Showcase diagram '$($diagram.name)' preview '$previewPath' is not listed in artifacts."
        $previewArtifact = $artifactsByPath[$previewPath]
        Assert-Condition -Condition ($previewArtifact.kind -in @('NativePreview', 'DesktopPreview', 'Preview')) -Message "Showcase diagram '$($diagram.name)' preview '$previewPath' points to a non-preview artifact."
        Assert-Condition -Condition ($preview.sha256 -eq $previewArtifact.sha256) -Message "Showcase diagram '$($diagram.name)' preview '$previewPath' hash does not match the artifact list."
        Assert-Condition -Condition ($preview.sizeBytes -eq $previewArtifact.sizeBytes) -Message "Showcase diagram '$($diagram.name)' preview '$previewPath' size does not match the artifact list."
        [void] $diagramPreviewPaths.Add($previewPath)
    }

    foreach ($proof in $diagramProofs) {
        $proofPath = $proof.relativePath
        Assert-Condition -Condition ($artifactsByPath.ContainsKey($proofPath)) -Message "Showcase diagram '$($diagram.name)' proof '$proofPath' is not listed in artifacts."
        $proofArtifact = $artifactsByPath[$proofPath]
        Assert-Condition -Condition ($proofArtifact.kind -in @('Inspection', 'StencilProfile', 'VisualQuality', 'Proof')) -Message "Showcase diagram '$($diagram.name)' proof '$proofPath' points to a non-proof artifact."
        Assert-Condition -Condition ($proof.sha256 -eq $proofArtifact.sha256) -Message "Showcase diagram '$($diagram.name)' proof '$proofPath' hash does not match the artifact list."
        Assert-Condition -Condition ($proof.sizeBytes -eq $proofArtifact.sizeBytes) -Message "Showcase diagram '$($diagram.name)' proof '$proofPath' size does not match the artifact list."
        if ($proof.kind -eq 'VisualQuality') {
            Assert-VisioShowcaseVisualQualityProof -RelativePath $proofPath -FullPath (Resolve-ShowcaseArtifactPath -RelativePath $proofPath)
        }
        [void] $diagramProofPaths.Add($proofPath)
    }
}

Assert-Condition -Condition ($diagramPackagePaths.Count -eq $packages.Count) -Message "Not every package artifact is referenced by a diagram. packages=$($packages.Count), referenced=$($diagramPackagePaths.Count)."
Assert-Condition -Condition ($diagramPreviewPaths.Count -eq $previews.Count) -Message "Not every preview artifact is referenced by a diagram. previews=$($previews.Count), referenced=$($diagramPreviewPaths.Count)."
Assert-Condition -Condition ($diagramProofPaths.Count -eq $proofs.Count) -Message "Not every structural proof artifact is referenced by a diagram. proofs=$($proofs.Count), referenced=$($diagramProofPaths.Count)."

$diagramCatalogLookup = @{}
$diagramCatalogToDiagrams = @{}
foreach ($diagram in $diagrams) {
    foreach ($catalog in (Get-RequiredArray -Value $diagram.proofSummary.stencilCatalogs)) {
        $catalogKey = $catalog.ToLowerInvariant()
        $diagramCatalogLookup[$catalogKey] = $catalog
        if (-not $diagramCatalogToDiagrams.ContainsKey($catalogKey)) {
            $diagramCatalogToDiagrams[$catalogKey] = @{}
        }

        $diagramCatalogToDiagrams[$catalogKey][$diagram.name] = $true
    }
}

Assert-Condition -Condition ($summaryCatalogLookup.Count -eq $diagramCatalogLookup.Count) -Message "Summary stencil catalog count does not match diagram proof summaries. summary=$($summaryCatalogLookup.Count), diagrams=$($diagramCatalogLookup.Count)."
foreach ($catalogKey in $diagramCatalogLookup.Keys) {
    Assert-Condition -Condition ($summaryCatalogLookup.ContainsKey($catalogKey)) -Message "Summary stencilCatalogs is missing diagram catalog '$($diagramCatalogLookup[$catalogKey])'."
    Assert-Condition -Condition ($coverageCatalogLookup.ContainsKey($catalogKey)) -Message "stencilCatalogCoverage is missing diagram catalog '$($diagramCatalogLookup[$catalogKey])'."

    $expectedDiagramNames = @($diagramCatalogToDiagrams[$catalogKey].Keys | Sort-Object)
    $actualDiagramNames = @(Get-RequiredArray -Value $coverageCatalogLookup[$catalogKey].diagrams | Sort-Object)
    Assert-Condition -Condition ($expectedDiagramNames.Count -eq $actualDiagramNames.Count) -Message "stencilCatalogCoverage diagram count mismatch for '$($diagramCatalogLookup[$catalogKey])'. expected=$($expectedDiagramNames.Count), actual=$($actualDiagramNames.Count)."
    for ($index = 0; $index -lt $expectedDiagramNames.Count; $index++) {
        Assert-Condition -Condition ($expectedDiagramNames[$index] -eq $actualDiagramNames[$index]) -Message "stencilCatalogCoverage diagram mismatch for '$($diagramCatalogLookup[$catalogKey])'. expected='$($expectedDiagramNames[$index])', actual='$($actualDiagramNames[$index])'."
    }
}

Assert-MarkdownSurface
Assert-GallerySurface

"Validated $($diagrams.Count) diagrams, $($packages.Count) packages, $($previews.Count) previews, $($proofs.Count) proofs, $($artifacts.Count) artifact hashes, proof totals, evidence totals, preview-format completeness, visual-quality proof artifacts, stencil proof summaries, stencil coverage, gallery review links, and Markdown fingerprints."
