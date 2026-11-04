BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Reader cmdlets' {
    It 'exposes built-in and modular Reader capabilities' {
        $capabilities = Get-OfficeDocumentCapability

        $capabilities.Id | Should -Contain 'officeimo.reader.word'
        $capabilities.Id | Should -Contain 'officeimo.reader.excel'
        $capabilities.Id | Should -Contain 'officeimo.reader.powerpoint'
        $capabilities.Id | Should -Contain 'officeimo.reader.markdown'
        $capabilities.Id | Should -Contain 'officeimo.reader.email'
        $capabilities.Id | Should -Contain 'officeimo.reader.email.store'
        $capabilities.Id | Should -Contain 'officeimo.reader.email.address-book'
        $capabilities.Id | Should -Contain 'officeimo.reader.pdf'
        $capabilities.Id | Should -Contain 'officeimo.reader.html'
        $capabilities.Id | Should -Contain 'officeimo.reader.csv'
        $capabilities.Id | Should -Contain 'officeimo.reader.json'
        $capabilities.Id | Should -Contain 'officeimo.reader.xml'
        $capabilities.Id | Should -Contain 'officeimo.reader.yaml'
        $capabilities.Id | Should -Contain 'officeimo.reader.zip'
        $capabilities.Id | Should -Contain 'officeimo.reader.epub'
        $capabilities.Id | Should -Contain 'officeimo.reader.visio'
        $capabilities.Id | Should -Contain 'officeimo.reader.rtf'
    }

    It 'accepts a caller-configured immutable Reader' {
        $handlerId = 'pswriteoffice.test.custom'
        $registrationType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Core' -TypeName 'OfficeIMO.Reader.ReaderHandlerRegistration' -CommandName 'Get-OfficeDocumentCapability'
        $builderType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Core' -TypeName 'OfficeIMO.Reader.OfficeDocumentReaderBuilder' -CommandName 'Get-OfficeDocumentCapability'
        $inputKindType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Core' -TypeName 'OfficeIMO.Reader.ReaderInputKind' -CommandName 'Get-OfficeDocumentCapability'
        $chunkType = Get-TestPSWriteOfficeType -AssemblyName 'OfficeIMO.Reader.Core' -TypeName 'OfficeIMO.Reader.ReaderChunk' -CommandName 'Get-OfficeDocumentCapability'

        $registration = [Activator]::CreateInstance($registrationType)
        $registration.Id = $handlerId
        $registration.DisplayName = 'Test custom Reader'
        $registration.Kind = [Enum]::Parse($inputKindType, 'Text')
        $registration.Extensions = [string[]]@('.custom')
        $readPathType = $registrationType.GetProperty('ReadPath').PropertyType
        $registration.ReadPath = [System.Management.Automation.LanguagePrimitives]::ConvertTo({
            param($Path, $Options, $CancellationToken)

            [Array]::CreateInstance($chunkType, 0)
        }.GetNewClosure(), $readPathType)

        $builder = [Activator]::CreateInstance($builderType)
        $builderType.GetMethod('AddHandler').Invoke($builder, @($registration, $false)) | Out-Null
        $reader = $builderType.GetMethod('Build').Invoke($builder, @())

        $capabilities = @(Get-OfficeDocumentCapability -Reader $reader -Origin Custom)
        $customCapability = $capabilities | Where-Object Id -EQ $handlerId
        $customCapability | Should -HaveCount 1
        $customCapability.Extensions | Should -Contain '.custom'
        $manifest = Get-OfficeDocumentCapability -Reader $reader -Origin Custom -Manifest
        $manifest.Handlers | Should -HaveCount 1
        $manifest.Handlers[0].Id | Should -Be $handlerId
        @(Get-OfficeDocumentCapability -Reader $reader -Origin OfficeIMO) | Should -HaveCount 0

        $customPath = Join-Path $TestDrive 'source.custom'
        Set-Content -Path $customPath -Value 'custom reader input' -Encoding UTF8

        { Read-OfficeDocument -Path $customPath -Reader $reader -NoHeadingChunks } |
            Should -Throw -ExpectedMessage '*immutable Reader*'
    }

    It 'reads Markdown files as chunks and a document envelope' {
        $path = Join-Path $TestDrive 'source.md'
        Set-Content -Path $path -Value "# Reader smoke`n`nOfficeIMO Reader keeps this text." -Encoding UTF8

        $chunks = @(Read-OfficeDocument -As Chunks -Path $path)
        $chunks.Count | Should -BeGreaterThan 0
        ($chunks.Text -join "`n") | Should -Match 'OfficeIMO Reader keeps this text'

        $document = Read-OfficeDocument -Path $path
        $document.Chunks.Count | Should -BeGreaterThan 0
        $document.Markdown | Should -Match 'Reader smoke'

        $json = Read-OfficeDocument -Path $path -As Json
        $json | Should -Match 'officeimo.document.read-result'
        $json | Should -Match 'OfficeIMO Reader keeps this text'

        $hierarchyOptions = New-OfficeDocumentHierarchyOptions -MaximumTokens 40 -OverlapTokens 0 -IncludeContextInText
        $hierarchy = $document | Read-OfficeDocument -As Hierarchy -HierarchyOptions $hierarchyOptions
        $hierarchy.Chunks.Count | Should -BeGreaterThan 0
        ($hierarchy.Chunks.Text -join "`n") | Should -Match 'OfficeIMO Reader keeps this text'
    }

    It 'reads HTML, JSON, and YAML files through modular Reader adapters' {
        $htmlPath = Join-Path $TestDrive 'source.html'
        Set-Content -Path $htmlPath -Value '<!doctype html><html><body><h1>Reader HTML</h1><p>OfficeIMO adapter text.</p></body></html>' -Encoding UTF8

        $htmlChunks = @(Read-OfficeDocument -As Chunks -Path $htmlPath)
        $htmlChunks.Count | Should -BeGreaterThan 0
        $htmlChunks[0].Kind.ToString() | Should -Be 'Html'
        ($htmlChunks.Text -join "`n") | Should -Match 'Reader HTML'
        ($htmlChunks.Text -join "`n") | Should -Match 'OfficeIMO adapter text'

        $jsonPath = Join-Path $TestDrive 'source.json'
        Set-Content -Path $jsonPath -Value '{"name":"Ada","score":42}' -Encoding UTF8

        $jsonChunks = @(Read-OfficeDocument -As Chunks -Path $jsonPath)
        $jsonChunks.Count | Should -BeGreaterThan 0
        $jsonChunks[0].Kind.ToString() | Should -Be 'Json'
        ($jsonChunks.Text -join "`n") | Should -Match '\$\.name'
        ($jsonChunks.Text -join "`n") | Should -Match 'Ada'

        $yamlPath = Join-Path $TestDrive 'source.yaml'
        Set-Content -Path $yamlPath -Value @(
            'service:'
            '  name: OfficeIMO'
            '  port: 443'
        ) -Encoding UTF8

        $yamlChunks = @(Read-OfficeDocument -As Chunks -Path $yamlPath)
        $yamlChunks.Count | Should -BeGreaterThan 0
        $yamlChunks[0].Kind.ToString() | Should -Be 'Yaml'
        ($yamlChunks.Text -join "`n") | Should -Match '\$\.service\.name'
        ($yamlChunks.Text -join "`n") | Should -Match 'OfficeIMO'
    }

    It 'reads RTF files through the semantic Reader adapter' {
        $path = Join-Path $TestDrive 'source.rtf'
        New-OfficeRtf -Path $path -Text 'Reader RTF adapter', 'Semantic chunk text' | Out-Null

        $chunks = @(Read-OfficeDocument -As Chunks -Path $path)
        $chunks.Count | Should -BeGreaterThan 0
        $chunks[0].Kind.ToString() | Should -Be 'Rtf'
        ($chunks.Text -join "`n") | Should -Match 'Reader RTF adapter'
        ($chunks.Text -join "`n") | Should -Match 'Semantic chunk text'
        ($chunks.Text -join "`n") | Should -Not -Match '\\rtf1'
    }

    It 'searches page-aware Reader results and projects page Markdown' {
        $path = Join-Path $TestDrive 'page-aware.rtf'
        [System.IO.File]::WriteAllText(
            $path,
            '{\rtf1\ansi First page text\page Second page retention period}',
            [System.Text.Encoding]::ASCII)

        $document = Read-OfficeDocument -Path $path -IncludePageLocations
        $document.Pages.Count | Should -Be 2

        $matches = @($document | Search-OfficeDocument -Pattern 'retention period' -SimpleMatch -WholeWord)
        $matches | Should -HaveCount 1
        $matches[0].Pages | Should -Be @(2)
        $matches[0].Match | Should -Be 'retention period'

        $pageMarkdown = @($document | Read-OfficeDocument -As Markdown -ByPage)
        $pageMarkdown | Should -HaveCount 2
        $pageMarkdown[1].Markdown | Should -Match 'Second page retention period'

        $combined = $document | Read-OfficeDocument -As Markdown
        $combined | Should -Match '<!-- page: 2/2; provenance: ExplicitBreak -->'
    }

    It 'searches Word Excel Markdown PST and OST together through PowerShell-native parameters' {
        $folder = Join-Path $TestDrive 'mixed-search'
        $nested = Join-Path $folder 'nested'
        New-Item -Path $nested -ItemType Directory | Out-Null

        $wordPath = Join-Path $folder 'policy.docx'
        New-OfficeWord -Path $wordPath {
            WordSection { WordParagraph -Text 'Synthetic Word evidence' }
        } | Out-Null

        $excelPath = Join-Path $folder 'register.xlsx'
        New-OfficeExcel -Path $excelPath {
            ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Synthetic Excel evidence' }
        } | Out-Null

        $markdownPath = Join-Path $nested 'notes.md'
        Set-Content -Path $markdownPath -Value '# Notes', 'Synthetic Markdown evidence' -Encoding UTF8
        $pstPath = Join-Path $folder 'mail.pst'
        $ostPath = Join-Path $nested 'offline.ost'
        [System.IO.File]::WriteAllBytes(
            $pstPath,
            [Convert]::FromBase64String((Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Assets\SyntheticMailStore.pst.b64') -Raw).Trim()))
        [System.IO.File]::WriteAllBytes(
            $ostPath,
            [Convert]::FromBase64String((Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Assets\SyntheticMailStore.ost.b64') -Raw).Trim()))
        Set-Content -Path (Join-Path $folder 'broken.docx') -Value 'not an OpenXML package' -Encoding UTF8
        Set-Content -Path (Join-Path $folder 'ignored.bin') -Value 'Synthetic unsupported input' -Encoding UTF8

        $readErrors = @()
        $matches = @(Search-OfficeDocument -Path $folder -Recurse -Pattern 'Synthetic' `
                -MaximumFiles 10 -MaximumStoreItems 10 -MaximumResults 10 -ThrottleLimit 2 `
                -ErrorVariable +readErrors -ErrorAction SilentlyContinue)

        ($matches.Path | Select-Object -Unique) | Should -HaveCount 5
        $matches.Path | Should -Contain $wordPath
        $matches.Path | Should -Contain $excelPath
        $matches.Path | Should -Contain $markdownPath
        $matches.Path | Should -Contain $pstPath
        $matches.Path | Should -Contain $ostPath
        $matches.DocumentType | Should -Contain 'Word'
        $matches.DocumentType | Should -Contain 'Excel'
        $matches.DocumentType | Should -Contain 'Markdown'
        ($matches | Where-Object DocumentType -EQ 'Email').Path | Select-Object -Unique | Should -HaveCount 2
        $matches.Match | Should -Not -Contain $null
        $matches.DocumentLimitReached | Should -Not -Contain $true
        $matches.SourceLimitReached | Should -Not -Contain $true
        $readErrors | Should -HaveCount 1

        $unlimited = @(Search-OfficeDocument -Path $markdownPath -Pattern 'Synthetic' `
                -MaximumFiles 0 -MaximumStoreItems 0 -MaximumResults 0)
        $unlimited | Should -HaveCount 1

        $store = Read-OfficeDocument -Path $pstPath -MaximumStoreItems 0
        $store.Kind.ToString() | Should -Be 'Email'
        $store.Metadata | Where-Object Name -EQ 'SelectionLimitReached' |
            Select-Object -ExpandProperty Value | Should -Not -Contain 'True'
    }

    It 'reports a configurable document ceiling without requiring collection objects' {
        $folder = Join-Path $TestDrive 'bounded-search'
        New-Item -Path $folder -ItemType Directory | Out-Null
        1..3 | ForEach-Object {
            Set-Content -Path (Join-Path $folder ("document-{0}.md" -f $_)) `
                -Value "Synthetic bounded document $_" -Encoding UTF8
        }

        $warnings = @()
        $matches = @(Search-OfficeDocument -Path $folder -Pattern 'Synthetic' -MaximumFiles 2 `
                -WarningVariable +warnings -WarningAction SilentlyContinue)

        ($matches.Path | Select-Object -Unique) | Should -HaveCount 2
        $matches.DocumentLimitReached | Should -Not -Contain $false
        ($warnings -join ' ') | Should -Match 'MaximumFiles \(2\)'
    }

    It 'applies extension filters before explicit paths consume the search ceiling' {
        $folder = Join-Path $TestDrive 'explicit-extension-filter'
        New-Item -Path $folder -ItemType Directory | Out-Null
        $excludedPath = Join-Path $folder 'first.json'
        $includedPath = Join-Path $folder 'second.md'
        Set-Content -Path $excludedPath -Value '{ "text": "Synthetic excluded document" }' -Encoding UTF8
        Set-Content -Path $includedPath -Value '# Synthetic included document' -Encoding UTF8

        $matches = @(Search-OfficeDocument -Path @($excludedPath, $includedPath) `
                -Extension md -Pattern 'Synthetic' -MaximumFiles 1)

        $matches | Should -HaveCount 1
        $matches[0].Path | Should -Be $includedPath
        $matches[0].DocumentLimitReached | Should -BeFalse
    }

    It 'applies extension filters to wildcard-expanded batch paths' {
        $folder = Join-Path $TestDrive 'wildcard-extension-filter'
        New-Item -Path $folder -ItemType Directory | Out-Null
        Set-Content -Path (Join-Path $folder 'first.json') `
            -Value '{ "text": "Synthetic excluded document" }' -Encoding UTF8
        $includedPath = Join-Path $folder 'second.md'
        Set-Content -Path $includedPath -Value '# Synthetic included document' -Encoding UTF8

        $documents = @(Read-OfficeDocument -Path (Join-Path $folder '*') `
                -Extension md -MaximumFiles 1)

        $documents | Should -HaveCount 1
        $documents[0].Source.Path | Should -Be $includedPath
    }

    It 'continues a batch after an individual document fails' {
        $folder = Join-Path $TestDrive 'resilient-batch'
        New-Item -Path $folder -ItemType Directory | Out-Null
        $goodPath = Join-Path $folder 'good.md'
        Set-Content -Path $goodPath -Value '# Good batch document' -Encoding UTF8
        Set-Content -Path (Join-Path $folder 'broken.docx') -Value 'not an OpenXML package' -Encoding UTF8

        $readErrors = @()
        $documents = @(Read-OfficeDocument -Path $folder -MaximumFiles 10 `
                -ThrottleLimit 2 `
                -ErrorVariable +readErrors -ErrorAction SilentlyContinue)

        $documents | Should -HaveCount 1
        $documents[0].Source.Path | Should -Be $goodPath
        $readErrors | Should -HaveCount 1
    }

    It 'creates configurable readers without requiring OfficeIMO option objects' {
        $runtime = New-TestTesseractExecutable -Directory $TestDrive
        $ocr = New-OfficeOcrOptions -LanguageExpression 'eng+pol' -TesseractPath $runtime -NoLanguageDownload
        $reader = New-OfficeDocumentReader -OcrOptions $ocr -MaximumStoreItems 2500 -MaximumConcurrentReads 2

        $reader | Should -Not -BeNullOrEmpty
        $reader.MaxConcurrentReads | Should -Be 2
        $reader.ProcessorPipeline.Count | Should -Be 1
    }

    It 'uses caller-selected search concurrency and rejects an undersized immutable Reader' {
        $path = Join-Path $TestDrive 'configured-search.md'
        Set-Content -Path $path -Value '# Concurrency', 'Synthetic configurable search' -Encoding UTF8
        $reader = New-OfficeDocumentReader -MaximumConcurrentReads 6

        $matches = @(Search-OfficeDocument -Path $path -Pattern 'Synthetic' `
                -Reader $reader -ThrottleLimit 6)

        $matches | Should -HaveCount 1
        $matches[0].Path | Should -Be $path
        {
            Search-OfficeDocument -Path $path -Pattern 'Synthetic' `
                -Reader (New-OfficeDocumentReader -MaximumConcurrentReads 2) `
                -ThrottleLimit 3 -ErrorAction Stop
        } | Should -Throw -ExpectedMessage '*requested batch concurrency (3)*Reader limit (2)*'
    }

    It 'reads Markdown tables and materializes deterministic table sidecars' {
        $path = Join-Path $TestDrive 'table.md'
        Set-Content -Path $path -Value @(
            '# Tables'
            ''
            '| Name | Score |'
            '| --- | ---: |'
            '| Ada | 42 |'
        ) -Encoding UTF8

        $tables = @(Read-OfficeDocument -As Tables -Path $path)
        $tables | Should -HaveCount 1
        $tables[0].Columns | Should -Be @('Name', 'Score')
        $tables[0].Rows[0][0] | Should -Be 'Ada'
        $tables[0].Rows[0][1] | Should -Be '42'

        $exports = @(Read-OfficeDocument -As Tables -Path $path -Bundle)
        $exports | Should -HaveCount 1
        $exports[0].Csv | Should -Match 'Name,Score'
        $exports[0].Markdown | Should -Match '\| Ada \| 42 \|'
        $exports[0].Json | Should -Match '"columns":\["Name","Score"\]'

        $outputDirectory = Join-Path $TestDrive 'table-exports'
        $materialized = @(Export-OfficeDocumentAsset -Type Table -Path $path -OutputDirectory $outputDirectory)
        $materialized | Should -HaveCount 3
        $materialized.Written | Should -Not -Contain $false
        Test-Path -LiteralPath (Join-Path $outputDirectory 'table-table-0000.csv') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $outputDirectory 'table-table-0000.md') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $outputDirectory 'table-table-0000.json') | Should -BeTrue
    }

    It 'reads and materializes embedded document assets' {
        $path = Join-Path $TestDrive 'assets.docx'
        $imagePath = Join-Path $PSScriptRoot 'Assets\CellImage.png'

        New-OfficeWord -Path $path {
            WordSection {
                WordParagraph { WordText 'Document with embedded asset' }
                WordParagraph { WordImage -Path $imagePath -Width 24 -Height 24 }
            }
        } | Out-Null

        $assets = @(Read-OfficeDocument -Path $path -As Assets | Where-Object Kind -EQ image)
        $assets.Count | Should -BeGreaterThan 0
        $assets[0].Kind | Should -Be 'image'
        $assets[0].MediaType | Should -Match 'image'
        $assets[0].PayloadBytes.Length | Should -BeGreaterThan 0

        $outputDirectory = Join-Path $TestDrive 'asset-exports'
        $materialized = @(Export-OfficeDocumentAsset -Path $path -Kind image -OutputDirectory $outputDirectory -ValidatePayloadHash)
        ($materialized | Where-Object Written).Count | Should -BeGreaterThan 0
        Test-Path -LiteralPath $materialized[0].Path | Should -BeTrue
    }

    It 'reads folders using extension filters' {
        $folder = Join-Path $TestDrive 'reader-folder'
        New-Item -Path $folder -ItemType Directory | Out-Null
        Set-Content -Path (Join-Path $folder 'first.md') -Value '# First' -Encoding UTF8
        Set-Content -Path (Join-Path $folder 'skip.txt') -Value 'skip me' -Encoding UTF8

        $chunks = @(Read-OfficeDocument -As Chunks -Path $folder -Extension md)
        ($chunks.Location.Path | Select-Object -Unique) | Should -HaveCount 1
        ($chunks.Text -join "`n") | Should -Match 'First'
        ($chunks.Text -join "`n") | Should -Not -Match 'skip me'
    }

    It 'summarizes folder ingestion for adapter-backed formats' {
        $folder = Join-Path $TestDrive 'ingest-folder'
        New-Item -Path $folder -ItemType Directory | Out-Null
        Set-Content -Path (Join-Path $folder 'page.html') -Value '<html><body><p>Indexed HTML</p></body></html>' -Encoding UTF8
        Set-Content -Path (Join-Path $folder 'data.json') -Value '{"item":"Indexed JSON"}' -Encoding UTF8
        Set-Content -Path (Join-Path $folder 'skip.txt') -Value 'skip me' -Encoding UTF8

        Set-Content -Path (Join-Path $folder 'config.yaml') -Value "name: Indexed YAML" -Encoding UTF8

        $result = Read-OfficeDocument -As Ingest -Path $folder -Extension html,json,yaml
        $result.FilesScanned | Should -Be 3
        $result.FilesParsed | Should -Be 3
        $result.ChunksProduced | Should -Be 3
        ($result.Chunks.Text -join "`n") | Should -Match 'Indexed HTML'
        ($result.Chunks.Text -join "`n") | Should -Match 'Indexed JSON'
        ($result.Chunks.Text -join "`n") | Should -Match 'Indexed YAML'
        ($result.Chunks.Text -join "`n") | Should -Not -Match 'skip me'
    }
}
