param(
    [string] $SiteRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Catalog.ps1')
. (Join-Path $PSScriptRoot 'PSWriteOfficeApiDocs.Documentation.ps1')
$apiRoot = Join-Path $SiteRoot 'data/apidocs/powershell'
$catalogPath = Join-Path $SiteRoot 'data/pswriteoffice_command_catalog.json'
$evidencePath = Join-Path $apiRoot 'documentation-manifest.json'
$manifest = Import-PowerShellDataFile -LiteralPath (Join-Path $apiRoot 'PSWriteOffice.psd1')
$source = @([pscustomobject]@{ Ref = (Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json).sourceRef })
$expected = @{
    SiteRootPath = $SiteRoot
    ExpectedSourceRef = $source[0].Ref
    ExpectedVersion = [version] $manifest.ModuleVersion
    ExpectedCommandCount = @($manifest.CmdletsToExport | Sort-Object -Unique).Count
}
$null = Assert-DocumentationCatalogSnapshot -CatalogPath $catalogPath -EvidencePath $evidencePath @expected

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-catalog-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporaryRoot | Out-Null
$temporaryCatalog = Join-Path $temporaryRoot 'catalog.json'
$temporaryEvidence = Join-Path $temporaryRoot 'documentation-manifest.json'
$cases = [Collections.Generic.List[string]]::new()
try {
    Copy-Item -LiteralPath $evidencePath -Destination $temporaryEvidence
    # The original regression preserved the version and both aggregate totals.
    foreach ($case in 'description', 'family-allocation', 'url', 'featured-command') {
        $catalog = Get-Content -LiteralPath $catalogPath -Raw | ConvertFrom-Json -Depth 20
        switch ($case) {
            'description' { $catalog.families[0].description = 'Stale description' }
            'family-allocation' { $catalog.families[0].commandCount--; $catalog.families[1].commandCount++ }
            'url' { $catalog.families[0].examplesUrl = 'https://example.invalid/stale' }
            'featured-command' { $catalog.families[0].featuredCommands = @('Get-StaleCommand') }
        }
        $catalog | ConvertTo-Json -Depth 20 -Compress | Set-Content -LiteralPath $temporaryCatalog -Encoding utf8
        if (-not (Test-DocumentationCatalog -Path $temporaryCatalog -ExpectedVersion $expected.ExpectedVersion -ExpectedCommandCount $expected.ExpectedCommandCount)) {
            throw "The '$case' fixture must retain the formerly accepted aggregate totals."
        }
        $rejected = $false
        try { $null = Assert-DocumentationCatalogSnapshot -CatalogPath $temporaryCatalog -EvidencePath $temporaryEvidence @expected }
        catch {
            if ($_.Exception.Message -notlike '*does not match its pinned source/derived evidence*') { throw }
            $rejected = $true
        }
        if (-not $rejected) { throw "An altered '$case' catalog was accepted." }
        $cases.Add($case)
    }

    # The same checked-in content remains valid in Windows and Linux checkouts.
    $content = [IO.File]::ReadAllText($catalogPath).Replace("`r`n", "`n")
    [IO.File]::WriteAllText($temporaryCatalog, $content.Replace("`n", "`r`n"), [Text.UTF8Encoding]::new($true))
    $null = Assert-DocumentationCatalogSnapshot -CatalogPath $temporaryCatalog -EvidencePath $temporaryEvidence @expected
    $cases.Add('bom-crlf-checkout')

    foreach ($case in 'wrong-source', 'missing-source-fingerprint') {
        $evidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json -Depth 10
        if ($case -eq 'wrong-source') { $evidence.sourceCommit = '0000000000000000000000000000000000000000' }
        else { $evidence.sourceCatalog = $null }
        $evidence | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $temporaryEvidence
        $rejected = $false
        try { $null = Assert-DocumentationCatalogSnapshot -CatalogPath $temporaryCatalog -EvidencePath $temporaryEvidence @expected }
        catch { $rejected = $true }
        if (-not $rejected) { throw "Invalid '$case' evidence was accepted." }
        $cases.Add($case)
    }

    # Completeness follows the pinned Git inventory, including sparse checkouts.
    $sourceFixture = Join-Path $temporaryRoot 'source'
    $sourceDocs = Join-Path $sourceFixture 'Website/content/project-docs/docs'
    New-Item -ItemType Directory -Path $sourceDocs -Force | Out-Null
    foreach ($page in '_index.md', 'word.md') {
        Set-Content -LiteralPath (Join-Path $sourceDocs $page) -Value '# Documentation fixture'
    }
    New-Item -ItemType Directory -Path (Join-Path $sourceFixture 'Build') | Out-Null
    Set-Content -LiteralPath (Join-Path $sourceFixture 'Build/Export-WebsiteDocumentationCatalog.ps1') -Value '# Producer fixture'
    & git -C $sourceFixture init --quiet
    & git -C $sourceFixture add -- 'Website/content/project-docs/docs' 'Build'
    & git -C $sourceFixture -c commit.gpgSign=false -c user.name='Documentation test' -c user.email='docs@example.invalid' commit --quiet -m 'Documentation fixture'
    if ($LASTEXITCODE -ne 0) { throw 'Could not commit the documentation fixture.' }
    $sourceCommit = & git -C $sourceFixture rev-parse HEAD
    if (-not (Test-CompleteSourceDocumentation -RepositoryRoot $sourceFixture -SourceCommit $sourceCommit)) {
        throw 'The complete tracked documentation fixture was rejected.'
    }
    & git -C $sourceFixture sparse-checkout set --no-cone '/Website/content/project-docs/docs/'
    if (Test-CompleteSourceDocumentation -RepositoryRoot $sourceFixture -SourceCommit $sourceCommit) {
        throw 'A source checkout without its catalog producer was accepted.'
    }
    $cases.Add('missing-catalog-producer')
    & git -C $sourceFixture sparse-checkout set --no-cone '/Website/content/project-docs/docs/_index.md' '/Build/'
    if ($LASTEXITCODE -ne 0 -or @(git -C $sourceFixture status --porcelain).Count -ne 0) {
        throw 'The partial documentation fixture must be a clean sparse checkout.'
    }
    if (Test-CompleteSourceDocumentation -RepositoryRoot $sourceFixture -SourceCommit $sourceCommit) {
        throw 'An incomplete sparse documentation checkout was accepted.'
    }
    $cases.Add('tracked-documentation-completeness')

    $siteFixture = Join-Path $temporaryRoot 'site'
    $fixtureDocs = Join-Path $siteFixture 'content/docs'
    New-Item -ItemType Directory -Path $fixtureDocs -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $SiteRoot 'content/docs/pswriteoffice') -Destination $fixtureDocs -Recurse
    Copy-Item -LiteralPath (Join-Path $SiteRoot 'content/docs/toc.json') -Destination (Join-Path $fixtureDocs 'toc.json')
    $pageEvidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json -Depth 15
    Assert-DocumentationPageSnapshot -SiteRootPath $siteFixture -Evidence $pageEvidence
    $firstPage = Join-Path $siteFixture $pageEvidence.documentation[0].sitePath
    $originalPage = [IO.File]::ReadAllBytes($firstPage)
    $fixtureToc = Join-Path $fixtureDocs 'toc.json'
    $originalToc = [IO.File]::ReadAllBytes($fixtureToc)
    foreach ($case in 'stale-page', 'missing-page', 'extra-page', 'stale-navigation') {
        switch ($case) {
            'stale-page' { Add-Content -LiteralPath $firstPage -Value 'Stale release syntax' }
            'missing-page' { Remove-Item -LiteralPath $firstPage }
            'extra-page' { Set-Content -LiteralPath (Join-Path $fixtureDocs 'pswriteoffice/extra.md') -Value '# Extra page' }
            'stale-navigation' {
                $toc = Get-Content -LiteralPath $fixtureToc -Raw | ConvertFrom-Json
                @($toc | Where-Object title -CEQ 'PSWriteOffice')[0].items[0].title = 'Stale navigation'
                $toc | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $fixtureToc
            }
        }
        $rejected = $false
        try { Assert-DocumentationPageSnapshot -SiteRootPath $siteFixture -Evidence $pageEvidence }
        catch { $rejected = $true }
        if (-not $rejected) { throw "The '$case' documentation snapshot was accepted." }
        [IO.File]::WriteAllBytes($firstPage, $originalPage)
        [IO.File]::WriteAllBytes($fixtureToc, $originalToc)
        $extra = Join-Path $fixtureDocs 'pswriteoffice/extra.md'
        if (Test-Path -LiteralPath $extra) { Remove-Item -LiteralPath $extra }
        $cases.Add($case)
    }

    # Existing navigation entries keep their order while adopting source titles.
    $navigationSource = Join-Path $temporaryRoot 'navigation-source'
    New-Item -ItemType Directory -Path $navigationSource | Out-Null
    Set-Content -LiteralPath (Join-Path $navigationSource 'word.md') -Value "---`ntitle: `"Automate Word Documents`"`n---`n# Word"
    $navigationToc = Join-Path $temporaryRoot 'toc.json'
    '[{"title":"PSWriteOffice","items":[{"title":"Word Cmdlets","path":"pswriteoffice/word/index.md"}]}]' | Set-Content -LiteralPath $navigationToc
    Sync-DocumentationToc -Source $navigationSource -TocPath $navigationToc
    $navigation = Get-Content -LiteralPath $navigationToc -Raw | ConvertFrom-Json
    if ($navigation[0].items[0].title -cne 'Automate Word Documents') { throw 'Existing navigation titles were not refreshed.' }
    $cases.Add('existing-navigation-title')

} finally {
    # Git creates hidden metadata and read-only loose objects on Windows.
    # This unique test directory contains only the fixture created above.
    foreach ($item in Get-ChildItem -LiteralPath $temporaryRoot -Force -Recurse) {
        $item.Attributes = $item.Attributes -band (-bnot ([IO.FileAttributes]::Hidden -bor [IO.FileAttributes]::ReadOnly))
    }
    # Directory.Delete also handles dot-prefixed Git metadata on Unix.
    [IO.Directory]::Delete($temporaryRoot, $true)
}

[PSCustomObject]@{ Status = 'passed'; CatalogCases = @($cases); CommandCount = $expected.ExpectedCommandCount }
