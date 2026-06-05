function Assert-EvidenceBoolean {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Evidence,

        [Parameter(Mandatory = $true)]
        [string] $PropertyName,

        [Parameter(Mandatory = $true)]
        [bool] $Expected,

        [Parameter(Mandatory = $true)]
        [string] $DiagramName
    )

    Assert-Condition -Condition ($Evidence.PSObject.Properties.Name -contains $PropertyName) -Message "Diagram '$DiagramName' evidence is missing '$PropertyName'."
    Assert-Condition -Condition ($Evidence.$PropertyName -is [bool]) -Message "Diagram '$DiagramName' evidence '$PropertyName' is not boolean."
    Assert-Condition -Condition ($Evidence.$PropertyName -eq $Expected) -Message "Diagram '$DiagramName' evidence '$PropertyName' mismatch. summary=$($Evidence.$PropertyName), expected=$Expected."
}

function Assert-EvidenceTotals {
    Assert-Condition -Condition ($summary.PSObject.Properties.Name -contains 'evidenceTotals') -Message 'Summary is missing evidenceTotals.'
    $totals = $summary.evidenceTotals
    foreach ($propertyName in @(
        'diagramCount',
        'nativeSvgPreviewDiagramCount',
        'nativePngPreviewDiagramCount',
        'completeNativePreviewDiagramCount',
        'desktopSvgPreviewDiagramCount',
        'desktopPngPreviewDiagramCount',
        'completeDesktopPreviewDiagramCount',
        'inspectionProofDiagramCount',
        'stencilProfileProofDiagramCount',
        'visualQualityProofDiagramCount',
        'cleanVisualQualityDiagramCount',
        'visualQualityIssueDiagramCount',
        'visualQualityIssueCount',
        'visualQualityErrorCount',
        'visualQualityWarningCount',
        'visualQualityInformationCount',
        'completeStructuralProofDiagramCount',
        'completeReviewProofDiagramCount',
        'completeNativeEvidenceDiagramCount',
        'completeDesktopEvidenceDiagramCount',
        'completePreviewEvidenceDiagramCount')) {
        Assert-Condition -Condition ($totals.PSObject.Properties.Name -contains $propertyName) -Message "evidenceTotals is missing '$propertyName'."
        Assert-Condition -Condition ($totals.$propertyName -is [int] -or $totals.$propertyName -is [long]) -Message "evidenceTotals.$propertyName is not numeric."
        Assert-Condition -Condition ([long] $totals.$propertyName -ge 0) -Message "evidenceTotals.$propertyName cannot be negative."
    }

    $expected = [ordered]@{
        diagramCount = [long] $diagrams.Count
        nativeSvgPreviewDiagramCount = 0L
        nativePngPreviewDiagramCount = 0L
        completeNativePreviewDiagramCount = 0L
        desktopSvgPreviewDiagramCount = 0L
        desktopPngPreviewDiagramCount = 0L
        completeDesktopPreviewDiagramCount = 0L
        inspectionProofDiagramCount = 0L
        stencilProfileProofDiagramCount = 0L
        visualQualityProofDiagramCount = 0L
        cleanVisualQualityDiagramCount = 0L
        visualQualityIssueDiagramCount = 0L
        visualQualityIssueCount = 0L
        visualQualityErrorCount = 0L
        visualQualityWarningCount = 0L
        visualQualityInformationCount = 0L
        completeStructuralProofDiagramCount = 0L
        completeReviewProofDiagramCount = 0L
        completeNativeEvidenceDiagramCount = 0L
        completeDesktopEvidenceDiagramCount = 0L
        completePreviewEvidenceDiagramCount = 0L
    }
    $missing = @{
        diagramsMissingNativeSvgPreview = New-Object System.Collections.Generic.List[string]
        diagramsMissingNativePngPreview = New-Object System.Collections.Generic.List[string]
        diagramsMissingCompleteNativePreview = New-Object System.Collections.Generic.List[string]
        diagramsMissingDesktopSvgPreview = New-Object System.Collections.Generic.List[string]
        diagramsMissingDesktopPngPreview = New-Object System.Collections.Generic.List[string]
        diagramsMissingCompleteDesktopPreview = New-Object System.Collections.Generic.List[string]
        diagramsMissingInspectionProof = New-Object System.Collections.Generic.List[string]
        diagramsMissingStencilProfileProof = New-Object System.Collections.Generic.List[string]
        diagramsMissingVisualQualityProof = New-Object System.Collections.Generic.List[string]
        diagramsWithVisualQualityIssues = New-Object System.Collections.Generic.List[string]
        diagramsMissingCompleteStructuralProof = New-Object System.Collections.Generic.List[string]
        diagramsMissingCompleteReviewProof = New-Object System.Collections.Generic.List[string]
        diagramsMissingCompleteNativeEvidence = New-Object System.Collections.Generic.List[string]
        diagramsMissingCompleteDesktopEvidence = New-Object System.Collections.Generic.List[string]
        diagramsMissingCompletePreviewEvidence = New-Object System.Collections.Generic.List[string]
    }

    foreach ($diagram in $diagrams) {
        Assert-Condition -Condition ($diagram.PSObject.Properties.Name -contains 'evidence') -Message "Showcase diagram '$($diagram.name)' is missing evidence."
        $evidence = $diagram.evidence
        $hasNativeSvg = Test-DiagramPreviewEvidence -Diagram $diagram -Kind 'NativePreview' -Format 'svg'
        $hasNativePng = Test-DiagramPreviewEvidence -Diagram $diagram -Kind 'NativePreview' -Format 'png'
        $hasDesktopSvg = Test-DiagramPreviewEvidence -Diagram $diagram -Kind 'DesktopPreview' -Format 'svg'
        $hasDesktopPng = Test-DiagramPreviewEvidence -Diagram $diagram -Kind 'DesktopPreview' -Format 'png'
        $hasInspection = Test-DiagramProofEvidence -Diagram $diagram -Kind 'Inspection'
        $hasStencilProfile = Test-DiagramProofEvidence -Diagram $diagram -Kind 'StencilProfile'
        $hasVisualQuality = Test-DiagramProofEvidence -Diagram $diagram -Kind 'VisualQuality'
        $visualQualitySummary = $null
        if ($hasVisualQuality) {
            $visualQualityProof = @($diagram.proofs | Where-Object { $_.kind -eq 'VisualQuality' })[0]
            $visualQualitySummary = Get-VisioShowcaseVisualQualityProofSummary -RelativePath $visualQualityProof.relativePath -FullPath (Resolve-ShowcaseArtifactPath -RelativePath $visualQualityProof.relativePath)
        }
        $hasCompleteNativePreview = $hasNativeSvg -and $hasNativePng
        $hasCompleteDesktopPreview = $hasDesktopSvg -and $hasDesktopPng
        $hasCompleteStructuralProof = $hasInspection -and $hasStencilProfile
        $hasCompleteReviewProof = $hasCompleteStructuralProof -and $hasVisualQuality
        $hasCompleteNativeEvidence = $hasCompleteNativePreview -and $hasCompleteReviewProof
        $hasCompleteDesktopEvidence = $hasCompleteDesktopPreview -and $hasCompleteReviewProof
        $hasCompletePreviewEvidence = ($hasCompleteNativePreview -or $hasCompleteDesktopPreview) -and $hasCompleteReviewProof

        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasNativeSvgPreview' -Expected $hasNativeSvg -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasNativePngPreview' -Expected $hasNativePng -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasDesktopSvgPreview' -Expected $hasDesktopSvg -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasDesktopPngPreview' -Expected $hasDesktopPng -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasInspectionProof' -Expected $hasInspection -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasStencilProfileProof' -Expected $hasStencilProfile -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasVisualQualityProof' -Expected $hasVisualQuality -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasCompleteNativePreview' -Expected $hasCompleteNativePreview -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasCompleteDesktopPreview' -Expected $hasCompleteDesktopPreview -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasCompleteStructuralProof' -Expected $hasCompleteStructuralProof -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasCompleteReviewProof' -Expected $hasCompleteReviewProof -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasCompleteNativeEvidence' -Expected $hasCompleteNativeEvidence -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasCompleteDesktopEvidence' -Expected $hasCompleteDesktopEvidence -DiagramName $diagram.name
        Assert-EvidenceBoolean -Evidence $evidence -PropertyName 'hasCompletePreviewEvidence' -Expected $hasCompletePreviewEvidence -DiagramName $diagram.name

        if ($hasNativeSvg) { $expected.nativeSvgPreviewDiagramCount++ } else { [void] $missing.diagramsMissingNativeSvgPreview.Add($diagram.name) }
        if ($hasNativePng) { $expected.nativePngPreviewDiagramCount++ } else { [void] $missing.diagramsMissingNativePngPreview.Add($diagram.name) }
        if ($hasCompleteNativePreview) { $expected.completeNativePreviewDiagramCount++ } else { [void] $missing.diagramsMissingCompleteNativePreview.Add($diagram.name) }
        if ($hasDesktopSvg) { $expected.desktopSvgPreviewDiagramCount++ } else { [void] $missing.diagramsMissingDesktopSvgPreview.Add($diagram.name) }
        if ($hasDesktopPng) { $expected.desktopPngPreviewDiagramCount++ } else { [void] $missing.diagramsMissingDesktopPngPreview.Add($diagram.name) }
        if ($hasCompleteDesktopPreview) { $expected.completeDesktopPreviewDiagramCount++ } else { [void] $missing.diagramsMissingCompleteDesktopPreview.Add($diagram.name) }
        if ($hasInspection) { $expected.inspectionProofDiagramCount++ } else { [void] $missing.diagramsMissingInspectionProof.Add($diagram.name) }
        if ($hasStencilProfile) { $expected.stencilProfileProofDiagramCount++ } else { [void] $missing.diagramsMissingStencilProfileProof.Add($diagram.name) }
        if ($hasVisualQuality) { $expected.visualQualityProofDiagramCount++ } else { [void] $missing.diagramsMissingVisualQualityProof.Add($diagram.name) }
        if ($null -ne $visualQualitySummary) {
            Assert-Condition -Condition ($diagram.PSObject.Properties.Name -contains 'visualQualitySummary') -Message "Showcase diagram '$($diagram.name)' is missing visualQualitySummary."
            Assert-Condition -Condition ($diagram.visualQualitySummary.hasProof -eq $true) -Message "Showcase diagram '$($diagram.name)' visualQualitySummary.hasProof mismatch."
            Assert-Condition -Condition ($diagram.visualQualitySummary.isClean -eq $visualQualitySummary.isClean) -Message "Showcase diagram '$($diagram.name)' visualQualitySummary.isClean mismatch."
            foreach ($propertyName in @('issueCount', 'errorCount', 'warningCount', 'informationCount')) {
                Assert-Condition -Condition ([long] $diagram.visualQualitySummary.$propertyName -eq [long] $visualQualitySummary.$propertyName) -Message "Showcase diagram '$($diagram.name)' visualQualitySummary.$propertyName mismatch."
            }

            $actualIssueKinds = @(Get-RequiredArray -Value $diagram.visualQualitySummary.issueKinds | Sort-Object)
            $expectedIssueKinds = @($visualQualitySummary.issueKinds)
            Assert-Condition -Condition ($actualIssueKinds.Count -eq $expectedIssueKinds.Count) -Message "Showcase diagram '$($diagram.name)' visualQualitySummary.issueKinds count mismatch."
            for ($index = 0; $index -lt $actualIssueKinds.Count; $index++) {
                Assert-Condition -Condition ($actualIssueKinds[$index] -eq $expectedIssueKinds[$index]) -Message "Showcase diagram '$($diagram.name)' visualQualitySummary.issueKinds mismatch."
            }

            if ($visualQualitySummary.isClean) { $expected.cleanVisualQualityDiagramCount++ }
            if ($visualQualitySummary.issueCount -gt 0) {
                $expected.visualQualityIssueDiagramCount++
                [void] $missing.diagramsWithVisualQualityIssues.Add($diagram.name)
            }

            $expected.visualQualityIssueCount += [long] $visualQualitySummary.issueCount
            $expected.visualQualityErrorCount += [long] $visualQualitySummary.errorCount
            $expected.visualQualityWarningCount += [long] $visualQualitySummary.warningCount
            $expected.visualQualityInformationCount += [long] $visualQualitySummary.informationCount
        }
        if ($hasCompleteStructuralProof) { $expected.completeStructuralProofDiagramCount++ } else { [void] $missing.diagramsMissingCompleteStructuralProof.Add($diagram.name) }
        if ($hasCompleteReviewProof) { $expected.completeReviewProofDiagramCount++ } else { [void] $missing.diagramsMissingCompleteReviewProof.Add($diagram.name) }
        if ($hasCompleteNativeEvidence) { $expected.completeNativeEvidenceDiagramCount++ } else { [void] $missing.diagramsMissingCompleteNativeEvidence.Add($diagram.name) }
        if ($hasCompleteDesktopEvidence) { $expected.completeDesktopEvidenceDiagramCount++ } else { [void] $missing.diagramsMissingCompleteDesktopEvidence.Add($diagram.name) }
        if ($hasCompletePreviewEvidence) { $expected.completePreviewEvidenceDiagramCount++ } else { [void] $missing.diagramsMissingCompletePreviewEvidence.Add($diagram.name) }
    }

    foreach ($propertyName in @($expected.Keys)) {
        Assert-Condition -Condition ([long] $totals.$propertyName -eq $expected[$propertyName]) -Message "evidenceTotals.$propertyName mismatch. summary=$($totals.$propertyName), diagrams=$($expected[$propertyName])."
    }

    foreach ($propertyName in @($missing.Keys)) {
        $actual = @(Get-RequiredArray -Value $totals.$propertyName | Sort-Object)
        $expectedMissing = @($missing[$propertyName].ToArray() | Sort-Object)
        Assert-Condition -Condition ($actual.Count -eq $expectedMissing.Count) -Message "evidenceTotals.$propertyName count mismatch. summary=$($actual.Count), diagrams=$($expectedMissing.Count)."
        for ($index = 0; $index -lt $actual.Count; $index++) {
            Assert-Condition -Condition ($actual[$index] -eq $expectedMissing[$index]) -Message "evidenceTotals.$propertyName mismatch. expected='$($expectedMissing[$index])', actual='$($actual[$index])'."
        }
    }

    if ($RequireNativePreviewFormatsPerDiagram) {
        Assert-Condition -Condition ($totals.nativeSvgPreviewDiagramCount -eq $diagrams.Count) -Message 'Not every diagram has a native SVG preview.'
        Assert-Condition -Condition ($totals.nativePngPreviewDiagramCount -eq $diagrams.Count) -Message 'Not every diagram has a native PNG preview.'
    }

    if ($RequireDesktopPreviewFormatsPerDiagram) {
        Assert-Condition -Condition ($totals.desktopSvgPreviewDiagramCount -eq $diagrams.Count) -Message 'Not every diagram has a desktop SVG preview.'
        Assert-Condition -Condition ($totals.desktopPngPreviewDiagramCount -eq $diagrams.Count) -Message 'Not every diagram has a desktop PNG preview.'
    }

    if ($RequireProofsPerDiagram) {
        Assert-Condition -Condition ($totals.inspectionProofDiagramCount -eq $diagrams.Count) -Message 'Not every diagram has an inspection proof.'
        Assert-Condition -Condition ($totals.stencilProfileProofDiagramCount -eq $diagrams.Count) -Message 'Not every diagram has a stencil-profile proof.'
        Assert-Condition -Condition ($totals.visualQualityProofDiagramCount -eq $diagrams.Count) -Message 'Not every diagram has a visual-quality proof.'
    }

    Assert-MarkdownContains -Needle "Complete native evidence: $($totals.completeNativeEvidenceDiagramCount)/$($totals.diagramCount)" -Message 'Markdown summary complete native evidence line is missing or stale.'
    Assert-MarkdownContains -Needle "Complete structural proof: $($totals.completeStructuralProofDiagramCount)/$($totals.diagramCount)" -Message 'Markdown summary complete structural proof line is missing or stale.'
    Assert-MarkdownContains -Needle "Complete review proof: $($totals.completeReviewProofDiagramCount)/$($totals.diagramCount)" -Message 'Markdown summary complete review proof line is missing or stale.'
    Assert-MarkdownContains -Needle "Clean visual quality: $($totals.cleanVisualQualityDiagramCount)/$($totals.diagramCount)" -Message 'Markdown summary clean visual-quality line is missing or stale.'
    Assert-MarkdownContains -Needle "Visual quality issues: $($totals.visualQualityIssueCount) ($($totals.visualQualityErrorCount) errors, $($totals.visualQualityWarningCount) warnings)" -Message 'Markdown summary visual-quality issue line is missing or stale.'
    Assert-MarkdownContains -Needle '## Evidence Coverage' -Message 'Markdown summary is missing the evidence coverage section.'
    Assert-MarkdownContains -Needle "Native SVG previews: $($totals.nativeSvgPreviewDiagramCount)/$($totals.diagramCount)" -Message 'Markdown summary native SVG evidence line is missing or stale.'
    Assert-MarkdownContains -Needle "Native PNG previews: $($totals.nativePngPreviewDiagramCount)/$($totals.diagramCount)" -Message 'Markdown summary native PNG evidence line is missing or stale.'
    Assert-MarkdownContains -Needle "Visual-quality proofs: $($totals.visualQualityProofDiagramCount)/$($totals.diagramCount)" -Message 'Markdown summary visual-quality evidence line is missing or stale.'
    Assert-MarkdownContains -Needle "Clean visual-quality proofs: $($totals.cleanVisualQualityDiagramCount)/$($totals.diagramCount)" -Message 'Markdown summary clean visual-quality evidence line is missing or stale.'
    Assert-GalleryContains -Needle 'href="#evidence-coverage"' -Message 'Gallery evidence coverage navigation link is missing.'
    Assert-GalleryContains -Needle '<h2 id="evidence-coverage">Evidence Coverage</h2>' -Message 'Gallery evidence coverage section is missing.'
    Assert-GalleryContains -Needle "<tr><td>Native SVG preview</td><td>$($totals.nativeSvgPreviewDiagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))/$($totals.diagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td><td>" -Message 'Gallery native SVG evidence row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Visual-quality proof</td><td>$($totals.visualQualityProofDiagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))/$($totals.diagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td><td>" -Message 'Gallery visual-quality evidence row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Clean visual quality</td><td>$($totals.cleanVisualQualityDiagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))/$($totals.diagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td><td>" -Message 'Gallery clean visual-quality evidence row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Complete review proof</td><td>$($totals.completeReviewProofDiagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))/$($totals.diagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td><td>" -Message 'Gallery complete review proof row is missing or stale.'
    Assert-GalleryContains -Needle "<tr><td>Complete native evidence</td><td>$($totals.completeNativeEvidenceDiagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))/$($totals.diagramCount.ToString('N0', [Globalization.CultureInfo]::InvariantCulture))</td><td>" -Message 'Gallery complete native evidence row is missing or stale.'
}
