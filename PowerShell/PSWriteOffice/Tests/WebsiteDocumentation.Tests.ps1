BeforeAll {
    $script:repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
    $script:catalogScript = Join-Path $script:repoRoot 'Build\Export-WebsiteDocumentationCatalog.ps1'
    $script:catalogPath = Join-Path $script:repoRoot 'WebsiteArtifacts\documentation\command-catalog.json'
    $script:projectManifestPath = Join-Path $script:repoRoot 'WebsiteArtifacts\project-manifest.json'
    $script:moduleManifestPath = Join-Path $script:repoRoot 'PSWriteOffice.psd1'
    $script:apiRoot = Join-Path $script:repoRoot 'WebsiteArtifacts\apidocs\powershell'
    $script:generatedHelpPath = Join-Path $script:repoRoot 'Docs\Generated\PSWriteOffice-help.xml'
    $script:sourceSnapshotManifestPath = Join-Path $script:apiRoot 'PSWriteOffice.psd1'
    $script:sourceExamplesRoot = Join-Path $script:repoRoot 'Examples'
    $script:snapshotExamplesRoot = Join-Path $script:apiRoot 'examples'
    $script:commandFamiliesGuidePath = Join-Path $script:repoRoot 'Website\content\project-docs\docs\command-families.md'
    $script:overviewGuidePath = Join-Path $script:repoRoot 'Website\content\project-docs\docs\overview.md'
    $script:projectDocsRoot = Join-Path $script:repoRoot 'Website\content\project-docs\docs'
    $script:projectDocsIndexPath = Join-Path $script:projectDocsRoot '_index.md'
}

Describe 'PSWriteOffice website documentation catalog' {
    It 'keeps the website example snapshot identical to the repository examples' {
        $sourceFiles = @(Get-ChildItem -LiteralPath $script:sourceExamplesRoot -Recurse -File)
        $snapshotFiles = @(Get-ChildItem -LiteralPath $script:snapshotExamplesRoot -Recurse -File)

        $sourceByPath = @{}
        foreach ($file in $sourceFiles) {
            $relative = $file.FullName.Substring($script:sourceExamplesRoot.Length).TrimStart('\', '/')
            $sourceByPath[$relative] = $file.FullName
        }
        $snapshotByPath = @{}
        foreach ($file in $snapshotFiles) {
            $relative = $file.FullName.Substring($script:snapshotExamplesRoot.Length).TrimStart('\', '/')
            $snapshotByPath[$relative] = $file.FullName
        }

        @($sourceByPath.Keys | Sort-Object) | Should -Be @($snapshotByPath.Keys | Sort-Object)
        foreach ($relative in $sourceByPath.Keys) {
            (Get-FileHash -LiteralPath $sourceByPath[$relative] -Algorithm SHA256).Hash |
                Should -Be (Get-FileHash -LiteralPath $snapshotByPath[$relative] -Algorithm SHA256).Hash -Because $relative
        }
    }

    It 'covers every exported cmdlet exactly once' {
        $outputPath = Join-Path $TestDrive 'command-catalog.json'
        & $script:catalogScript -RepositoryRoot $script:repoRoot -OutputPath $outputPath | Out-Null

        $module = Import-PowerShellDataFile -LiteralPath $script:moduleManifestPath
        $catalog = Get-Content -LiteralPath $outputPath -Raw | ConvertFrom-Json

        $catalog.module.commandCount | Should -Be @($module.CmdletsToExport).Count
        $catalog.module.aliasCount | Should -Be @($module.AliasesToExport).Count
        $catalog.module.familyCount | Should -Be @($catalog.families).Count
        (@($catalog.families | Measure-Object -Property commandCount -Sum).Sum) | Should -Be $catalog.module.commandCount
        @($catalog.families | Where-Object commandCount -LT 1).Count | Should -Be 0
    }

    It 'keeps the committed catalog deterministic and current' {
        $outputPath = Join-Path $TestDrive 'command-catalog.json'
        & $script:catalogScript `
            -RepositoryRoot $script:repoRoot `
            -OutputPath $outputPath `
            -ManifestPath $script:sourceSnapshotManifestPath | Out-Null

        (Get-Content -LiteralPath $outputPath -Raw).Trim() |
            Should -Be (Get-Content -LiteralPath $script:catalogPath -Raw).Trim()
    }

    It 'keeps guide family counts aligned with the generated catalog' {
        $catalog = Get-Content -LiteralPath $script:catalogPath -Raw | ConvertFrom-Json
        $expected = @{}
        foreach ($family in $catalog.families) {
            $expected[[string] $family.title] = [int] $family.commandCount
        }

        $tableRows = [regex]::Matches(
            (Get-Content -LiteralPath $script:commandFamiliesGuidePath -Raw),
            '(?m)^\|\s*(?<title>[^|]+?)\s*\|\s*(?<count>\d+)\s*\|')
        $tableRows.Count | Should -Be $expected.Count
        foreach ($row in $tableRows) {
            $title = $row.Groups['title'].Value.Trim()
            $expected.ContainsKey($title) | Should -BeTrue -Because "'$title' should be a generated command family"
            [int] $row.Groups['count'].Value | Should -Be $expected[$title] -Because "'$title' should match the generated catalog"
        }

        $overviewRows = [regex]::Matches(
            (Get-Content -LiteralPath $script:overviewGuidePath -Raw),
            '(?m)^-\s+\*\*(?<title>.+?)\s+—\s+(?<count>\d+)\s+commands:')
        foreach ($row in $overviewRows) {
            $title = $row.Groups['title'].Value.Trim()
            $expected.ContainsKey($title) | Should -BeTrue -Because "'$title' should be a generated command family"
            [int] $row.Groups['count'].Value | Should -Be $expected[$title] -Because "'$title' should match the generated catalog"
        }
    }

    It 'accepts a filename-only catalog output path' {
        $fileName = 'command-catalog-' + [guid]::NewGuid().ToString('N') + '.json'
        $expectedPath = Join-Path $TestDrive $fileName
        Test-Path -LiteralPath $expectedPath | Should -BeFalse
        Push-Location $TestDrive
        try {
            $result = & $script:catalogScript `
                -RepositoryRoot $script:repoRoot `
                -OutputPath $fileName

            $result.OutputPath | Should -Be $expectedPath
            Test-Path -LiteralPath $expectedPath | Should -BeTrue
        } finally {
            Pop-Location
        }
    }

    It 'publishes real docs, examples, and API surfaces at the source snapshot version' {
        $sourceSnapshot = Import-PowerShellDataFile -LiteralPath $script:sourceSnapshotManifestPath
        $catalog = Get-Content -LiteralPath $script:catalogPath -Raw | ConvertFrom-Json
        $project = Get-Content -LiteralPath $script:projectManifestPath -Raw | ConvertFrom-Json

        $project.version | Should -Be ([string] $sourceSnapshot.ModuleVersion)
        $catalog.module.version | Should -Be ([string] $sourceSnapshot.ModuleVersion)
        $project.surfaces.docs | Should -BeTrue
        $project.surfaces.examples | Should -BeTrue
        $project.surfaces.apiPowerShell | Should -BeTrue
        $project.artifacts.docs | Should -Be 'Website/content/project-docs'
        Test-Path -LiteralPath (Join-Path $script:repoRoot $project.artifacts.docs) | Should -BeTrue
        Test-Path -LiteralPath (Join-Path (Join-Path $script:repoRoot $project.artifacts.docs) 'docs\_index.md') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:repoRoot $project.artifacts.examples) | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:repoRoot $project.artifacts.documentationCatalog) | Should -BeTrue
    }

    It 'keeps the committed PowerShell API bundle aligned with every exported cmdlet' {
        $module = Import-PowerShellDataFile -LiteralPath $script:moduleManifestPath
        $expected = @($module.CmdletsToExport) | Sort-Object -Unique
        $apiManifest = Import-PowerShellDataFile -LiteralPath (Join-Path $script:apiRoot 'PSWriteOffice.psd1')
        [xml] $help = Get-Content -LiteralPath (Join-Path $script:apiRoot 'PSWriteOffice-help.xml') -Raw
        $metadata = Get-Content -LiteralPath (Join-Path $script:apiRoot 'command-metadata.json') -Raw | ConvertFrom-Json

        $manifestCommands = @($apiManifest.CmdletsToExport) | Sort-Object -Unique
        $helpCommands = @($help.helpItems.command | ForEach-Object { [string] $_.details.name }) | Sort-Object -Unique
        $metadataCommands = @($metadata.commands | ForEach-Object { [string] $_.name }) | Sort-Object -Unique
        $expectedAliases = @($module.AliasesToExport) | Sort-Object -Unique
        $metadataAliases = @($metadata.commands.aliases) | Sort-Object -Unique

        $manifestCommands | Should -Be $expected
        $helpCommands | Should -Be $expected
        $metadataCommands | Should -Be $expected
        $metadataAliases | Should -Be $expectedAliases
        @($metadata.commands | Where-Object { -not $_.sourcePath }).Count | Should -Be 0

        $exportCsv = $metadata.commands | Where-Object name -EQ 'Export-OfficeCsv'
        $exportCsv.sourcePath | Should -Be 'Sources/PSWriteOffice/Cmdlets/Csv/ExportOfficeCsvCommand.cs'
        (Get-Content -LiteralPath (Join-Path $script:repoRoot $exportCsv.sourcePath))[$exportCsv.sourceLine - 1] |
            Should -Match '\bclass\s+ExportOfficeCsvCommand\b'
    }

    It 'keeps the generated module help synchronized with the API bundle' {
        $apiHelpPath = Join-Path $script:apiRoot 'PSWriteOffice-help.xml'

        Test-Path -LiteralPath $script:generatedHelpPath | Should -BeTrue
        (Get-FileHash -LiteralPath $script:generatedHelpPath -Algorithm SHA256).Hash |
            Should -Be (Get-FileHash -LiteralPath $apiHelpPath -Algorithm SHA256).Hash
    }

    It 'publishes discoverable comparison and legacy migration guides against current commands' {
        $module = Import-PowerShellDataFile -LiteralPath $script:moduleManifestPath
        $exportedCommands = @($module.CmdletsToExport)
        $guideSlugs = @(
            'compare-importexcel-excelfast'
            'compare-office-automation-options'
            'migrate-from-legacy-modules'
            'migrate-from-pswriteword'
            'migrate-from-pswriteexcel'
            'migrate-from-pswritepdf'
        )
        $index = Get-Content -LiteralPath $script:projectDocsIndexPath -Raw

        foreach ($slug in $guideSlugs) {
            $path = Join-Path $script:projectDocsRoot "$slug.md"
            Test-Path -LiteralPath $path | Should -BeTrue
            $content = Get-Content -LiteralPath $path -Raw
            $description = [regex]::Match($content, '(?m)^description:\s*"(?<value>[^"]+)"$').Groups['value'].Value

            $description.Length | Should -BeGreaterOrEqual 120 -Because "$slug needs a useful search summary"
            $description.Length | Should -BeLessOrEqual 160 -Because "$slug should avoid routine search-result truncation"
            $content | Should -Not -Match '(?m)^#\s+' -Because 'the page title already renders the only H1'
            $index | Should -Match ([regex]::Escape("/docs/pswriteoffice/$slug/"))
        }

        foreach ($command in @(
            'New-OfficeWord'
            'Join-OfficeWordDocument'
            'New-OfficeExcel'
            'Import-OfficeExcel'
            'New-OfficePdf'
            'Join-OfficePdf'
            'Split-OfficePdf'
            'Get-OfficePdfText'
            'Set-OfficePdfForm'
            'ConvertFrom-OfficePdfHtml'
        )) {
            $exportedCommands | Should -Contain $command -Because 'migration mappings must point at an exported command'
        }
    }
}
