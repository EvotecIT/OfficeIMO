BeforeAll {
    $manifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) { $env:PSWRITEOFFICE_MODULE_MANIFEST } else { Join-Path $PSScriptRoot '..\PSWriteOffice.psd1' }
    Import-Module $manifest -Global -ErrorAction Stop
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
}

Describe 'Consolidated document reading' {
    It 'reuses a normalized result for text, JSON, chunks, structure and hierarchy without reopening its source' {
        $path = Join-Path $TestDrive 'read-once.md'
        Set-Content -LiteralPath $path -Value "# Evidence`n`nDelivery is ready." -Encoding UTF8
        $document = Read-OfficeDocument -LiteralPath $path
        Remove-Item -LiteralPath $path
        $document | Read-OfficeDocument -As Text | Should -Match 'Delivery is ready'
        $document | Read-OfficeDocument -As Markdown -NoPageMarkers | Should -Match '# Evidence'
        ($document | Read-OfficeDocument -As Json | ConvertFrom-Json).SchemaId | Should -Not -BeNullOrEmpty
        @($document | Read-OfficeDocument -As Chunks).Count | Should -BeGreaterThan 0
        ($document | Read-OfficeDocument -As Structure).GetType().Name | Should -Be 'OfficeDocumentStructuredExtractionResult'
        ($document | Read-OfficeDocument -As Hierarchy).GetType().Name | Should -Be 'ReaderChunkHierarchyResult'
    }

    It 'binds literal bracketed FileInfo paths and reports missing inputs independently' {
        $path = Join-Path $TestDrive 'read[1].md'
        Set-Content -LiteralPath $path -Value 'Retained evidence' -Encoding UTF8
        Get-Item -LiteralPath $path | Read-OfficeDocument -As Text | Should -Match 'Retained evidence'
        $failures = @()
        $documents = @(Read-OfficeDocument -LiteralPath @((Join-Path $TestDrive 'missing.md'), $path) -ErrorAction SilentlyContinue -ErrorVariable +failures)
        $documents | Should -HaveCount 1
        $failures | Should -HaveCount 1
    }

    It 'selects folders explicitly, recurses on request and reports selection ceilings' {
        $folder = Join-Path $TestDrive 'documents'
        $child = Join-Path $folder 'child'
        $null = New-Item -ItemType Directory -Path $child -Force
        Set-Content -LiteralPath (Join-Path $folder 'first.md') -Value 'First'
        Set-Content -LiteralPath (Join-Path $child 'second.md') -Value 'Second'
        @(Read-OfficeDocument $folder -Extension md).Count | Should -Be 1
        @(Read-OfficeDocument $folder -Recurse -Extension md -MaximumFiles 0).Count | Should -Be 2
        $profile = New-OfficeDocumentReaderOptions -MaximumTraversalEntries 100
        $profile.FolderOptions.Extensions = [string[]] @('md')
        $configured = New-OfficeDocumentReader -Options $profile
        @(Read-OfficeDocument $folder -Reader $configured).Count | Should -Be 1
        @(Search-OfficeDocument $folder -Reader $configured -Pattern 'First').Count | Should -Be 1
        $warnings = @()
        @(Read-OfficeDocument $folder -Recurse -Extension md -MaximumFiles 1 -WarningAction SilentlyContinue -WarningVariable +warnings).Count | Should -Be 1
        $warnings | Should -Not -BeNullOrEmpty
    }

    It 'routes nested format options and lets flat parameters override a caller-owned snapshot' {
        $path = Join-Path $TestDrive 'options.md'
        Set-Content -LiteralPath $path -Value '# Evidence'
        $options = New-OfficeDocumentReaderOptions
        $options.ComputeHashes = $true
        $read = Read-OfficeDocument $path -Options $options, @{ Markdown = @{ ChunkByHeadings = $false } } -NoHashes
        $read.Source.SourceHash | Should -BeNullOrEmpty
        $options.ComputeHashes | Should -BeTrue
        { Read-OfficeDocument $path -Options @{ Markdown = @{ UnrecognisedField = $true } } -ErrorAction Stop } | Should -Throw '*Unknown or read-only*'
    }

    It 'distinguishes worksheet positions from numeric names and retains only the selected sheet' {
        $path = Join-Path $TestDrive 'sheets.xlsx'
        $workbook = [OfficeIMO.Excel.ExcelDocument]::Create()
        try {
            $null = $workbook.AddWorksheet('2').CellValue(1, 1, 'Named sheet evidence')
            $null = $workbook.AddWorksheet('Evidence').CellValue(1, 1, 'Second sheet evidence')
            $workbook.Save($path)
        } finally { $workbook.Dispose() }
        Read-OfficeDocument $path -Sheet 2 -NoHeaderRow -As Text | Should -Match 'Second sheet evidence'
        Read-OfficeDocument $path -Sheet '2' -NoHeaderRow -As Text | Should -Match 'Named sheet evidence'
        $pages = @(Read-OfficeDocument $path -Sheet 2 -NoHeaderRow -As Text -ByPage)
        $pages | Should -HaveCount 1
        $pages[0].Page.Name | Should -Be 'Evidence'
        $pages[0].Text | Should -Match 'Second sheet evidence'
    }

    It 'keeps contiguous text across bounded chunks and emits worksheet values once in either text view' {
        $textPath = Join-Path $TestDrive 'contiguous.txt'
        $source = ('a' * 8001) + "`n" + ('b' * 8001)
        [IO.File]::WriteAllText($textPath, $source, [Text.UTF8Encoding]::new($false))
        Read-OfficeDocument $textPath -As Text | Should -BeExactly $source
        $path = Join-Path $TestDrive 'table-once.xlsx'
        $workbook = [OfficeIMO.Excel.ExcelDocument]::Create()
        try {
            $sheet = $workbook.AddWorksheet('Evidence')
            $null = $sheet.CellValue(1, 1, 'Item')
            $null = $sheet.CellValue(2, 1, 'Delivery')
            $workbook.Save($path)
        } finally { $workbook.Dispose() }
        $text = Read-OfficeDocument $path -As Text
        [regex]::Matches($text, 'Delivery').Count | Should -Be 1
        $pages = @(Read-OfficeDocument $path -As Text -ByPage)
        $pages | Should -HaveCount 1
        $pages[0].Text | Should -BeExactly $text
    }

    It 'returns flat table models and in-memory bundles without writing sidecars' {
        $folder = Join-Path $TestDrive 'in-memory-table'
        $null = New-Item -ItemType Directory -Path $folder
        $path = Join-Path $folder 'table.md'
        Set-Content -LiteralPath $path -Value "| Item | Score |`n| --- | --- |`n| Delivery | 42 |" -Encoding UTF8
        $document = Read-OfficeDocument $path
        $document | Read-OfficeDocument -As Text | Should -Match 'Delivery'
        $tables = @($document | Read-OfficeDocument -As Tables)
        $tables | Should -HaveCount 1
        $tables[0].Rows[0][0] | Should -Be 'Delivery'
        $bundles = @($document | Read-OfficeDocument -As Tables -Bundle -Indented)
        $bundles | Should -HaveCount 1
        $bundles[0].GetType().Name | Should -Be 'ReaderTableExportBundle'
        @(Get-ChildItem -LiteralPath $folder).Count | Should -Be 1
    }

    It 'reports ingestion outcomes and counts chunks even when their payloads are omitted' {
        $first = Join-Path $TestDrive 'ingest-a.md'
        $second = Join-Path $TestDrive 'ingest-b.md'
        Set-Content -LiteralPath $first -Value 'First evidence'
        Set-Content -LiteralPath $second -Value 'Second evidence'
        $summary = Read-OfficeDocument -Path $first, $second -As Ingest -NoChunks
        $summary.FilesParsed | Should -Be 2
        $summary.FilesScanned | Should -Be 2
        $summary.ChunksProduced | Should -BeGreaterThan 0
        @($summary.Chunks).Count | Should -Be 0
        $summary.BytesRead | Should -BeGreaterThan 0
    }

    It 'accounts total bytes through native batch admission and includes rejected inputs in ingestion status' {
        $large = Join-Path $TestDrive 'large.md'
        $first = Join-Path $TestDrive 'bounded-first.md'
        $second = Join-Path $TestDrive 'bounded-second.md'
        [IO.File]::WriteAllText($large, ('large ' * 20), [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText($first, 'aa', [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText($second, 'bbb', [Text.UTF8Encoding]::new($false))
        $failures = @()
        $summary = Read-OfficeDocument -LiteralPath $large, $first, $second -As Ingest -MaximumTotalBytes 5 -ErrorAction SilentlyContinue -ErrorVariable +failures
        $summary.FilesScanned | Should -Be 3
        $summary.FilesParsed | Should -Be 2
        $summary.FilesSkipped | Should -Be 1
        $summary.BytesRead | Should -Be 5
        $failures | Should -HaveCount 1
    }

    It 'keeps incompatible view switches and immutable Reader overrides explicit' {
        $path = Join-Path $TestDrive 'reader.md'
        Set-Content -LiteralPath $path -Value 'Evidence'
        $reader = New-OfficeDocumentReader -MaximumConcurrentReads 4
        Read-OfficeDocument $path -Reader $reader -As Text | Should -Match 'Evidence'
        { Read-OfficeDocument $path -Reader $reader -NoFootnotes -ErrorAction Stop } | Should -Throw '*immutable Reader*'
        { Read-OfficeDocument $path -As Document -ByPage -ErrorAction Stop } | Should -Throw '*ByPage*'
        { Read-OfficeDocument $path -As Text -Bundle -ErrorAction Stop } | Should -Throw '*Bundle*'
        $document = Read-OfficeDocument $path
        { $document | Read-OfficeDocument -As Ocr -ErrorAction Stop } | Should -Throw '*original image or PDF*'
    }

    It 'captures reusable reader defaults and applies flat overrides without changing the reader' {
        $path = Join-Path $TestDrive 'configured-defaults.txt'
        [IO.File]::WriteAllText($path, 'Configured evidence', [Text.UTF8Encoding]::new($false))
        $options = New-OfficeDocumentReaderOptions -Projection TextAndMetadata
        $options.MaxInputBytes = 128
        $options.ComputeHashes = $false
        $reader = New-OfficeDocumentReader -Options $options -MaximumConcurrentReads 4
        $options.MaxInputBytes = 1
        $copy = $reader.DefaultOptions
        $copy.MaxInputBytes = 1
        $document = Read-OfficeDocument $path -Reader $reader
        $document.Source.SourceHash | Should -BeNullOrEmpty
        $document | Read-OfficeDocument -As Text | Should -Match 'Configured evidence'
        @(Search-OfficeDocument $path -Reader $reader -Pattern 'Configured').Count | Should -Be 1
        { Read-OfficeDocument $path -Reader $reader -MaximumInputBytes 1 -ErrorAction Stop } | Should -Throw '*MaxInputBytes*'
        $reader.DefaultOptions.MaxInputBytes | Should -Be 128
        [IO.File]::WriteAllText($path, ('large ' * 100), [Text.UTF8Encoding]::new($false))
        { Search-OfficeDocument $path -Reader $reader -Pattern 'large' -ErrorAction Stop } | Should -Throw '*MaxInputBytes*'
    }

    It 'binds familiar encoding names and carries them through the native text reader' {
        $path = Join-Path $TestDrive 'latin.txt'
        [IO.File]::WriteAllBytes($path, [byte[]] @(99, 97, 102, 233))
        $options = New-OfficeDocumentReaderOptions -Encoding latin1
        Read-OfficeDocument $path -As Text -Options $options | Should -Be ('caf' + [char] 0xE9)
        (New-OfficeDocumentReaderOptions -Encoding utf8).TextEncoding.WebName | Should -Be 'utf-8'
        { New-OfficeDocumentReaderOptions -Encoding 'not-a-real-encoding' -ErrorAction Stop } | Should -Throw '*encoding*'
    }

    It 'uses the configured folder traversal policy for both reading and searching' {
        $folder = Join-Path $TestDrive 'limited-walk'
        $null = New-Item -ItemType Directory -Path $folder
        Set-Content -LiteralPath (Join-Path $folder 'a.md') -Value 'Evidence'
        Set-Content -LiteralPath (Join-Path $folder 'b.md') -Value 'Evidence'
        $reader = New-OfficeDocumentReader -Options (New-OfficeDocumentReaderOptions -MaximumTraversalEntries 1)
        { Read-OfficeDocument $folder -Reader $reader -ErrorAction Stop } | Should -Throw '*MaxTraversalEntries*'
        { Search-OfficeDocument $folder -Reader $reader -Pattern 'Evidence' -ErrorAction Stop } | Should -Throw '*MaxTraversalEntries*'
    }

    It 'creates OCR options without executing a provider and enriches ordinary text only when requested' {
        $marker = Join-Path $TestDrive 'recognition-started.txt'
        $runtime = New-TestTesseractExecutable -Directory $TestDrive -CreateFilePath $marker
        $ocr = New-OfficeOcrOptions -TesseractPath $runtime -Language English -NoLanguageDownload `
            -SessionOptions @{ Engine = @{ Dpi = 96; Timeout = '00:01:00' } } -Dpi 300 -TimeoutSeconds 10
        Test-Path -LiteralPath $marker | Should -BeFalse
        $image = Join-Path $PSScriptRoot 'Assets\CellImage.png'
        $null = Read-OfficeDocument $image
        Test-Path -LiteralPath $marker | Should -BeFalse
        Read-OfficeDocument $image -As Text -OcrOptions $ocr | Should -Match 'OfficeIMO OCR'
        Test-Path -LiteralPath $marker | Should -BeTrue
        $review = Read-OfficeDocument $image -As Ocr -OcrOptions $ocr
        $review.GetType().Name | Should -Be 'PdfSearchableOcrReview'
        @($review.Ocr.Pages).Count | Should -BeGreaterThan 0
    }
}
