function Get-ShowcaseCardUrls {
    param($Card)
    @($Card.downloads.url) + @($Card.source_url, $Card.image) + @($Card.previews.url) | Where-Object { $_ } | Sort-Object -Unique
}

function Get-ShowcaseSourceBindings {
    param([string] $RepositoryRoot, $Catalog, $Card)
    $urls = @(Get-ShowcaseCardUrls $Card)
    $paths = @($Card.source_path) + @($Catalog.artifacts | Where-Object {
        $_.sourceRoot -ceq 'repository' -and ('/downloads/showcase/' + $_.destination) -cin $urls
    } | ForEach-Object source)
    foreach ($path in $paths | Sort-Object -Unique) {
        [ordered]@{ path = $path; sha256 = (Get-FileHash -LiteralPath (Resolve-ShowcasePath $RepositoryRoot $path) -Algorithm SHA256).Hash.ToLowerInvariant() }
    }
}

function Get-ShowcaseEvidenceProducerBinding {
    param([string] $RepositoryRoot)
    $paths = @('Website/scripts/Invoke-ShowcasePowerShellExample.ps1', 'Website/scripts/ShowcaseEvidence.Helpers.ps1', 'Website/scripts/ShowcaseEvidence.Provenance.ps1', 'Website/scripts/Build-ShowcaseEvidence.ps1')
    $files = @(foreach ($path in $paths) {
        [ordered]@{ path = $path; sha256 = (Get-FileHash -LiteralPath (Resolve-ShowcasePath $RepositoryRoot $path) -Algorithm SHA256).Hash.ToLowerInvariant() }
    })
    [ordered]@{ path = $files[0].path; sha256 = $files[0].sha256; dependencies = @($files | Select-Object -Skip 1) }
}

function Get-ShowcaseGenerationReceiptHash {
    param($Generation)
    # Hash the immutable producer claim after preview rendering. Builder-owned download
    # bindings and the later human review are validated separately and are not producer data.
    $claim = [ordered]@{}
    foreach ($name in 'exampleId', 'state', 'startedUtc', 'completedUtc', 'host', 'sources', 'evidenceProducer', 'package', 'conversions', 'outputs', 'previews') {
        $claim[$name] = $Generation.$name
    }
    foreach ($name in 'startedUtc', 'completedUtc') {
        $claim[$name] = ([DateTimeOffset] $Generation.$name).ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($claim | ConvertTo-Json -Depth 30 -Compress))
    [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)).ToLowerInvariant()
}

function Test-ShowcaseHashBindings {
    param([object[]] $Actual, [object[]] $Expected)
    if ($Actual.Count -ne $Expected.Count) { return $false }
    foreach ($binding in $Expected) {
        if (@($Actual | Where-Object { $_.path -ceq $binding.path -and $_.sha256 -ceq $binding.sha256 }).Count -ne 1) { return $false }
    }
    return $true
}

function Test-ShowcaseGenerationReceipt {
    param($Generation, $Card, $Catalog, [hashtable] $CurrentHashes)
    try {
        if ($Generation.exampleId -cne $Card.id -or
            -not $Generation.startedUtc -or -not $Generation.completedUtc -or
            [DateTimeOffset] $Generation.startedUtc -gt [DateTimeOffset] $Generation.completedUtc -or
            -not $Generation.host.powershell -or -not $Generation.host.os -or -not $Generation.host.architecture -or
            $Generation.package.name -cne 'PSWriteOffice' -or
            -not ($Generation.package.files -is [array]) -or @($Generation.package.files).Count -lt 1 -or
            @($Generation.package.files).Count -ne $Generation.package.fileCount) { return $false }
        $files = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($file in $Generation.package.files) {
            if (-not $file.path -or $file.path -match '(^/|\\|(^|/)\.\.(/|$)|:)' -or
                $file.sha256 -cnotmatch '^[a-f0-9]{64}$' -or -not $files.Add($file.path)) { return $false }
        }
        if (-not $files.Contains('PSWriteOffice.psd1')) { return $false }
        $pdfs = @($Generation.outputs | Where-Object { $_.path.EndsWith('.pdf', [StringComparison]::Ordinal) })
        if ($pdfs.Count -ne 1 -or @($Generation.conversions).Count -ne 1) { return $false }
        $conversion = $Generation.conversions[0]
        if ($conversion.outputPath -cne $pdfs[0].path -or $conversion.outputSha256 -cne $pdfs[0].sha256 -or
            $conversion.routeId -cne 'docx-pdf' -or $conversion.command -cne 'Export-OfficeDocumentPdf' -or
            -not $conversion.routeEvidence -or -not $conversion.report -or
            -not ($conversion.report.Diagnostics -is [array]) -or -not ($conversion.warnings -is [array]) -or
            $pdfs[0].diagnostics.pageCount -lt 1 -or -not ($pdfs[0].diagnostics.fonts -is [array])) { return $false }
        $urls = @($Card.previews.url)
        $previewArtifacts = @($Catalog.artifacts | Where-Object { $_.previewSource -and ('/downloads/showcase/' + $_.destination) -cin $urls })
        if ($urls.Count -lt 1 -or @($Generation.previews).Count -ne $urls.Count -or $previewArtifacts.Count -ne $urls.Count) { return $false }
        foreach ($artifact in $previewArtifacts) {
            $path = '/downloads/showcase/' + $artifact.destination
            $preview = @($Generation.previews | Where-Object path -CEQ $path)
            $input = @($Catalog.artifacts | Where-Object { $_.sourceRoot -ceq 'documents' -and -not $_.previewSource -and $_.source -ceq $artifact.previewSource })
            if ($preview.Count -ne 1 -or $input.Count -ne 1) { return $false }
            $rendering = $preview[0].rendering
            if ($rendering.tool -cne 'pdftocairo' -or -not $rendering.version -or
                $rendering.executableSha256 -cnotmatch '^[a-f0-9]{64}$' -or
                $rendering.page -ne $artifact.previewPage -or
                $rendering.inputSha256 -cne $CurrentHashes['/downloads/showcase/' + $input[0].destination] -or
                $rendering.sha256 -cne $CurrentHashes[$path] -or
                -not ($rendering.diagnostics -is [array])) { return $false }
        }
        return $Generation.receiptSha256 -cmatch '^[a-f0-9]{64}$' -and
            $Generation.receiptSha256 -ceq (Get-ShowcaseGenerationReceiptHash $Generation)
    } catch { return $false }
}

function Complete-ShowcaseGenerationEvidence {
    param([string] $RepositoryRoot, $Catalog, [object[]] $Artifacts, [object[]] $Previous, [hashtable] $Generated, [string[]] $ReviewedExampleId, [string] $ModuleVersion)
    $currentHashes = @{}
    foreach ($artifact in $Artifacts) { $currentHashes[$artifact.path] = $artifact.sha256 }
    $producer = Get-ShowcaseEvidenceProducerBinding $RepositoryRoot
    foreach ($id in $ReviewedExampleId) {
        if ($id -cnotin @($Catalog.cards | Where-Object source_language -CEQ 'powershell' | ForEach-Object id)) {
            throw "Preview review must name a PowerShell showcase example: $id"
        }
        if ($Generated.ContainsKey($id)) { throw "Record '$id' preview review in a separate manifest-only invocation after inspecting the generated preview." }
    }
    foreach ($card in $Catalog.cards | Where-Object source_language -CEQ 'powershell') {
        $previousMatches = @($Previous | Where-Object exampleId -CEQ $card.id)
        $generation = if ($Generated.ContainsKey($card.id)) { $Generated[$card.id] } else { $previousMatches | Select-Object -First 1 }
        $urls = @(Get-ShowcaseCardUrls $card)
        $bindings = @($Artifacts | Where-Object { $_.path -cin $urls } | ForEach-Object { [ordered]@{ path = $_.path; sha256 = $_.sha256 } })
        $valid = $null -ne $generation -and $generation.state -ceq 'generated' -and $generation.package.version -ceq $ModuleVersion
        if (-not $Generated.ContainsKey($card.id) -and $previousMatches.Count -ne 1) { $valid = $false }
        if ($valid) {
            if ($generation.evidenceProducer.path -cne $producer.path -or $generation.evidenceProducer.sha256 -cne $producer.sha256 -or @($generation.evidenceProducer.dependencies).Count -ne $producer.dependencies.Count) { $valid = $false }
            foreach ($dependency in $producer.dependencies) {
                if (@($generation.evidenceProducer.dependencies | Where-Object { $_.path -ceq $dependency.path -and $_.sha256 -ceq $dependency.sha256 }).Count -ne 1) { $valid = $false }
            }
            $sourceBindings = @(Get-ShowcaseSourceBindings -RepositoryRoot $RepositoryRoot -Catalog $Catalog -Card $card)
            if ($sourceBindings.Count -ne @($generation.sources).Count) { $valid = $false }
            foreach ($source in $sourceBindings) {
                if (@($generation.sources | Where-Object { $_.path -ceq $source.path -and $_.sha256 -ceq $source.sha256 }).Count -ne 1) { $valid = $false }
            }
            $expectedOutputs = @($Catalog.artifacts | Where-Object {
                $_.sourceRoot -ceq 'documents' -and -not $_.previewSource -and ('/downloads/showcase/' + $_.destination) -cin $urls
            } | ForEach-Object { '/downloads/showcase/' + $_.destination })
            if ($expectedOutputs.Count -eq 0 -or @($generation.outputs).Count -ne $expectedOutputs.Count) { $valid = $false }
            foreach ($path in $expectedOutputs) {
                $output = @($generation.outputs | Where-Object path -CEQ $path)
                if ($output.Count -ne 1 -or -not $currentHashes.ContainsKey($path) -or $currentHashes[$path] -cne $output[0].sha256) { $valid = $false }
            }
            if (-not (Test-ShowcaseGenerationReceipt -Generation $generation -Card $card -Catalog $Catalog -CurrentHashes $currentHashes)) { $valid = $false }
            # Fresh receipts bind previews below; retained receipts must still bind the whole download set.
            if (-not $Generated.ContainsKey($card.id)) {
                if (@($generation.artifacts).Count -ne $bindings.Count) { $valid = $false }
                foreach ($binding in $bindings) {
                    if (@($generation.artifacts | Where-Object { $_.path -ceq $binding.path -and $_.sha256 -ceq $binding.sha256 }).Count -ne 1) { $valid = $false }
                }
            }
        }
        if (-not $valid) {
            if ($card.id -cin $ReviewedExampleId) { throw "Generate '$($card.id)' before recording its preview review; generation evidence is missing or stale." }
            [ordered]@{ exampleId = $card.id; state = $(if ($generation) { 'invalidated' } else { 'not-recorded' }); previewReview = [ordered]@{ state = 'pending' }; artifacts = $bindings }
            continue
        }
        $generation | Add-Member -NotePropertyName artifacts -NotePropertyValue $bindings -Force
        if ($Generated.ContainsKey($card.id) -or -not $generation.previewReview) {
            $generation | Add-Member -NotePropertyName previewReview -NotePropertyValue ([ordered]@{ state = 'pending' }) -Force
        }
        if ($generation.previewReview.state -ceq 'reviewed') {
            $previewBindings = @($bindings | Where-Object { $_.path -cin @($card.previews.url) })
            $reviewIsCurrent = Test-ShowcaseHashBindings -Actual @($generation.previewReview.artifacts) -Expected $previewBindings
            try {
                $reviewIsCurrent = $reviewIsCurrent -and $generation.previewReview.reviewedUtc -and
                    [DateTimeOffset] $generation.previewReview.reviewedUtc -ge [DateTimeOffset] $generation.completedUtc
            } catch { $reviewIsCurrent = $false }
            if (-not $reviewIsCurrent) {
                $generation.previewReview = [ordered]@{ state = 'pending' }
            }
        }
        if ($card.id -cin $ReviewedExampleId) {
            $generation.previewReview = [ordered]@{ state = 'reviewed'; reviewedUtc = [datetime]::UtcNow.ToString('o'); artifacts = @($bindings | Where-Object { $_.path -cin @($card.previews.url) }) }
        }
        $generation
    }
}
