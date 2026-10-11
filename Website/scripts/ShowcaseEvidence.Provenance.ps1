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

function Complete-ShowcaseGenerationEvidence {
    param([string] $RepositoryRoot, $Catalog, [object[]] $Artifacts, [object[]] $Previous, [hashtable] $Generated, [string[]] $ReviewedExampleId, [string] $ModuleVersion)
    $currentHashes = @{}
    foreach ($artifact in $Artifacts) { $currentHashes[$artifact.path] = $artifact.sha256 }
    foreach ($id in $ReviewedExampleId) {
        if ($id -cnotin @($Catalog.cards | Where-Object source_language -CEQ 'powershell' | ForEach-Object id)) {
            throw "Preview review must name a PowerShell showcase example: $id"
        }
    }
    foreach ($card in $Catalog.cards | Where-Object source_language -CEQ 'powershell') {
        $generation = if ($Generated.ContainsKey($card.id)) { $Generated[$card.id] } else { @($Previous | Where-Object exampleId -CEQ $card.id) | Select-Object -First 1 }
        $urls = @(Get-ShowcaseCardUrls $card)
        $bindings = @($Artifacts | Where-Object { $_.path -cin $urls } | ForEach-Object { [ordered]@{ path = $_.path; sha256 = $_.sha256 } })
        $valid = $null -ne $generation -and $generation.state -ceq 'generated' -and $generation.package.version -ceq $ModuleVersion
        if ($valid) {
            $sourceBindings = @(Get-ShowcaseSourceBindings -RepositoryRoot $RepositoryRoot -Catalog $Catalog -Card $card)
            $valid = $sourceBindings.Count -eq @($generation.sources).Count
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
        if ($card.id -cin $ReviewedExampleId) {
            $generation.previewReview = [ordered]@{ state = 'reviewed'; reviewedUtc = [datetime]::UtcNow.ToString('o'); artifacts = @($bindings | Where-Object { $_.path -cin @($card.previews.url) }) }
        }
        $generation
    }
}
