BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        $sourceRoot = Join-Path (Join-Path (Join-Path $PSScriptRoot '..') 'Sources') 'PSWriteOffice'

        if (-not $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES) {
            $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
        }

        if (-not $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION) {
            $releasePath = Join-Path (Join-Path $sourceRoot 'bin') 'Release'
            $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = if (Test-Path $releasePath) { 'Release' } else { 'Debug' }
        }

        Join-Path (Join-Path $PSScriptRoot '..') 'PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'CSV cmdlets' {
    It 'exposes NoHeader instead of bool header toggles on CSV commands' {
        (Get-Command ConvertTo-OfficeCsv).Parameters.Keys | Should -Contain 'NoHeader'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Contain 'NoHeader'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Contain 'NoHeader'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'NoHeader'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'AsDataTable'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'AsDataReader'
        (Get-Command ConvertFrom-OfficeCsv).Parameters.Mode.ParameterType.FullName | Should -Be 'PSWriteOffice.Cmdlets.Csv.CsvReadMode'
        (Get-Command Import-OfficeCsv).Parameters.Mode.ParameterType.FullName | Should -Be 'PSWriteOffice.Cmdlets.Csv.CsvReadMode'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Not -Contain 'Mode'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Contain 'CompressionType'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Contain 'CompressionType'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'CompressionType'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Contain 'MaxDecompressedBytes'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'MaxDecompressedBytes'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Contain 'ParseErrorAction'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'ParseErrorAction'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Contain 'ProgressInterval'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'ProgressInterval'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'InferSchema'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'ColumnType'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'DuplicateHeaderBehavior'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'NullValue'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'DateTimeFormats'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'QuoteParsingMode'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Contain 'StaticColumns'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Contain 'NullValue'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Contain 'DateTimeFormat'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Contain 'UseUtc'

        (Get-Command ConvertTo-OfficeCsv).Parameters.Keys | Should -Not -Contain 'IncludeHeader'
        (Get-Command ConvertTo-OfficeCsv).Parameters.Keys | Should -Not -Contain 'CompressionType'
        (Get-Command ConvertFrom-OfficeCsv).Parameters.Keys | Should -Not -Contain 'CompressionType'
        (Get-Command Export-OfficeCsv).Parameters.Keys | Should -Not -Contain 'IncludeHeader'
        (Get-Command Get-OfficeCsv).Parameters.Keys | Should -Not -Contain 'HasHeaderRow'
        (Get-Command Import-OfficeCsv).Parameters.Keys | Should -Not -Contain 'HasHeaderRow'
    }

    It 'exposes CSV row import with idiomatic command names' {
        (Get-Command Import-OfficeCsv).CommandType | Should -Be 'Cmdlet'
        (Get-Command ConvertFrom-OfficeCsv).CommandType | Should -Be 'Cmdlet'
        { Get-Command Get-OfficeCsvData -ErrorAction Stop } | Should -Throw
    }

    It 'converts objects to CSV and reads them back' {
        $rows = @(
            [pscustomobject]@{ Region = 'NA'; Revenue = 100 }
            [pscustomobject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        $csvText = @($rows | ConvertTo-OfficeCsv)
        $csvText.Count | Should -Be 3
        $csvText[0] | Should -Be 'Region,Revenue'
        $csvText[1] | Should -Be 'NA,100'
        $csvText[2] | Should -Be 'EMEA,200'

        $path = Join-Path $TestDrive 'data.csv'
        $rows | Export-OfficeCsv -Path $path | Out-Null

        Test-Path $path | Should -BeTrue

        $data = Import-OfficeCsv -Path $path
        $data.Count | Should -Be 2
        $data[0].Region | Should -Be 'NA'
    }

    It 'imports quoted multiline and padded fields as row objects' {
        $path = Join-Path $TestDrive 'multiline-padded.csv'
        [System.IO.File]::WriteAllText(
            $path,
            "Name,Note,Value`nAlpha,`"one`ntwo`",1`nBeta,plain")

        $rows = @(Import-OfficeCsv -Path $path)

        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'Alpha'
        $rows[0].Note | Should -Be "one`ntwo"
        $rows[1].Name | Should -Be 'Beta'
        $rows[1].Value | Should -Be ''
    }

    It 'uses configured null tokens and date formats when converting and exporting rows' {
        $rows = @(
            [pscustomobject]@{
                Name = 'Alpha'
                Created = [datetime]::SpecifyKind([datetime]'2026-07-07T13:45:00', [System.DateTimeKind]::Utc)
                Value = $null
            }
            [pscustomobject]@{
                Name = 'Beta'
                Created = [datetime]::SpecifyKind([datetime]'2026-07-07T14:15:00', [System.DateTimeKind]::Utc)
                Value = $null
            }
        )

        $csvText = @($rows | ConvertTo-OfficeCsv -NullValue '<null>' -DateTimeFormat 'yyyyMMdd-HHmm' -UseUtc)
        $csvText[1] | Should -Be 'Alpha,20260707-1345,<null>'
        $csvText[2] | Should -Be 'Beta,20260707-1415,<null>'

        $path = Join-Path $TestDrive 'formatted.csv'
        $rows | Export-OfficeCsv -Path $path -NullValue '<null>' -DateTimeFormat 'yyyyMMdd-HHmm' -UseUtc

        (Get-Content -LiteralPath $path)[1] | Should -Be 'Alpha,20260707-1345,<null>'
        (Get-Content -LiteralPath $path)[2] | Should -Be 'Beta,20260707-1415,<null>'
    }

    It 'imports duplicate headers, null tokens, static columns, and strict quote settings through OfficeIMO options' {
        $path = Join-Path $TestDrive 'parity.csv'
        Set-Content -LiteralPath $path -Value "Name,Name,Value`nAlpha,Beta,<null>" -Encoding UTF8

        $row = Import-OfficeCsv -Path $path -NullValue '<null>' -StaticColumns @{ SourceFile = 'parity.csv' }

        $row.Name | Should -Be 'Alpha'
        $row.Name_2 | Should -Be 'Beta'
        $row.Value | Should -BeNullOrEmpty
        $row.SourceFile | Should -Be 'parity.csv'

        { Import-OfficeCsv -Path $path -DuplicateHeaderBehavior Throw -ErrorAction Stop } |
            Should -Throw '*duplicate*'

        { ConvertFrom-OfficeCsv -Text "Name,Value`nAlpha,`"one`"two" -QuoteParsingMode Strict -ErrorAction Stop } |
            Should -Throw '*quoted*'
    }

    It 'round-trips compressed CSV files through PSWriteOffice cmdlets' {
        $path = Join-Path $TestDrive 'compressed.csv.gz'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType GZip

        Test-Path -LiteralPath $path | Should -BeTrue

        $row = Import-OfficeCsv -Path $path -CompressionType GZip
        $row.Name | Should -Be 'Alpha'
        $row.Value | Should -Be '1'
    }

    It 'allows append to create a compressed CSV file when the target has no content' {
        $path = Join-Path $TestDrive 'append-create.csv.gz'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType GZip -Append

        $row = Import-OfficeCsv -Path $path -CompressionType GZip
        $row.Name | Should -Be 'Alpha'
        $row.Value | Should -Be '1'
    }

    It 'keeps load-time options on file import parameter sets' {
        $parameterSets = (Get-Command Import-OfficeCsv).Parameters['NullValue'].ParameterSets.Keys

        $parameterSets | Should -Contain 'PathDelimiter'
        $parameterSets | Should -Contain 'LiteralPathDelimiter'
        $parameterSets | Should -Not -Contain 'Document'
    }

    It 'imports CSV rows as a DataTable for database workflows' {
        $path = Join-Path $TestDrive 'datatable.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBeta,2" -Encoding UTF8

        $table = Import-OfficeCsv -Path $path -AsDataTable

        $table.GetType() | Should -Be ([System.Data.DataTable])
        $table.TableName | Should -Be 'datatable'
        $table.Columns.Count | Should -Be 2
        $table.Columns['Name'].DataType | Should -Be ([string])
        $table.Rows.Count | Should -Be 2
        $table.Rows[0]['Name'] | Should -Be 'Alpha'
        $table.Rows[1]['Value'] | Should -Be '2'
    }

    It 'preserves normalized empty DataTable fields' {
        $path = Join-Path $TestDrive 'datatable-missing.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBeta" -Encoding UTF8

        $table = Import-OfficeCsv -Path $path -AsDataTable

        $table.Rows[1]['Value'] | Should -Be ''
    }

    It 'imports header-only CSV files as empty DataTables' {
        $path = Join-Path $TestDrive 'datatable-header-only.csv'
        Set-Content -LiteralPath $path -Value 'Name,Value' -Encoding UTF8

        $table = Import-OfficeCsv -Path $path -AsDataTable

        $table.GetType() | Should -Be ([System.Data.DataTable])
        $table.Columns.Count | Should -Be 2
        $table.Columns['Name'].DataType | Should -Be ([string])
        $table.Rows.Count | Should -Be 0
    }

    It 'exports DataTable input directly as CSV rows' {
        $path = Join-Path $TestDrive 'datatable-export.csv'
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)
        [void] $table.Rows.Add('Beta', 2)

        Export-OfficeCsv -InputObject $table -Path $path

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'exports DataView input directly as CSV rows' {
        $path = Join-Path $TestDrive 'dataview-export.csv'
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)
        [void] $table.Rows.Add('Alpha', 1)
        $view = [System.Data.DataView]::new($table)
        $view.Sort = 'Name ASC'

        Export-OfficeCsv -InputObject $view -Path $path

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'exports IDataReader input directly as CSV rows' {
        $path = Join-Path $TestDrive 'reader-export.csv'
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)
        [void] $table.Rows.Add('Beta', 2)
        $reader = $table.CreateDataReader()
        try {
            Export-OfficeCsv -InputObject $reader -Path $path
        } finally {
            $reader.Dispose()
        }

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'exports and imports GZip compressed CSV files' {
        $path = Join-Path $TestDrive 'compressed.csv.gz'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $rows | Export-OfficeCsv -Path $path -CompressionType GZip -CompressionLevel Fastest

        $bytes = [System.IO.File]::ReadAllBytes($path)
        $bytes[0] | Should -Be 0x1F
        $bytes[1] | Should -Be 0x8B

        $imported = @(Import-OfficeCsv -Path $path -CompressionType GZip)
        $document = Get-OfficeCsv -Path $path -CompressionType GZip
        $table = Import-OfficeCsv -Path $path -CompressionType GZip -AsDataTable

        $imported.Count | Should -Be 2
        $imported[1].Name | Should -Be 'Beta'
        $document.Header | Should -Be @('Name', 'Value')
        $table.Rows.Count | Should -Be 2
        $table.Rows[0]['Value'] | Should -Be '1'
    }

    It 'infers GZip compression from export paths by default' {
        $path = Join-Path $TestDrive 'compressed-inferred.csv.gz'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        $imported = @(Import-OfficeCsv -Path $path)

        $imported.Count | Should -Be 1
        $imported[0].Name | Should -Be 'Alpha'
    }

    It 'exports DataTable input as GZip compressed CSV files' {
        $path = Join-Path $TestDrive 'datatable-compressed.csv.gz'
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)

        Export-OfficeCsv -InputObject $table -Path $path -CompressionType GZip

        $imported = @(Import-OfficeCsv -Path $path -CompressionType GZip)
        $imported.Count | Should -Be 1
        $imported[0].Name | Should -Be 'Alpha'
    }

    It 'exports and imports Deflate compressed CSV files' {
        $path = Join-Path $TestDrive 'compressed.csv.deflate'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType Deflate

        $imported = @(Import-OfficeCsv -Path $path -CompressionType Deflate)

        $imported.Count | Should -Be 1
        $imported[0].Name | Should -Be 'Alpha'
    }

    It 'passes parser safety and normalization options through import surfaces' {
        $path = Join-Path $TestDrive 'smart-quotes.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,$([char]0x201C)Hello$([char]0x201D)" -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path -NormalizeQuotes -MaxFieldLength 16 -InternStrings)

        $data[0].Value | Should -Be '"Hello"'
    }

    It 'can collect and skip malformed CSV rows' {
        $path = Join-Path $TestDrive 'malformed.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBroken,`"one`"two`nBeta,2" -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path -QuoteParsingMode Strict -ParseErrorAction SkipRow -CollectParseErrors -ErrorAction SilentlyContinue)

        $data.Name | Should -Be @('Alpha', 'Beta')
    }

    It 'reports collected parse errors before returning streamed data readers' {
        $path = Join-Path $TestDrive 'malformed-reader.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBroken,`"one`"two`nBeta,2" -Encoding UTF8

        $parseErrors = $null
        $reader = Import-OfficeCsv -Path $path -AsDataReader -Mode Stream -QuoteParsingMode Strict -ParseErrorAction SkipRow -CollectParseErrors -ErrorAction SilentlyContinue -ErrorVariable parseErrors
        try {
            $reader.GetType().Name | Should -Be 'CsvDataReader'
            $rows = 0
            while ($reader.Read()) {
                $rows++
            }

            $rows | Should -Be 2
        } finally {
            $reader.Dispose()
        }

        $parseErrors.Count | Should -BeGreaterThan 0
    }

    It 'reports collected parse errors before returning streamed CSV documents' {
        $path = Join-Path $TestDrive 'malformed-document.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBroken,`"one`"two`nBeta,2" -Encoding UTF8

        $parseErrors = $null
        $document = Get-OfficeCsv -Path $path -QuoteParsingMode Strict -ParseErrorAction SkipRow -CollectParseErrors -ErrorAction SilentlyContinue -ErrorVariable parseErrors

        @($document.AsEnumerable()).Count | Should -Be 2
        $parseErrors.Count | Should -BeGreaterThan 0
    }

    It 'infers DataTable schema when requested' {
        $path = Join-Path $TestDrive 'datatable-infer.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBeta,2" -Encoding UTF8

        $table = Import-OfficeCsv -Path $path -AsDataTable -InferSchema

        $table.Columns['Value'].DataType | Should -Be ([int])
        $table.Rows[1]['Value'] | Should -Be 2
    }

    It 'imports CSV rows as a typed IDataReader for database workflows' {
        $path = Join-Path $TestDrive 'reader-import.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBeta,2" -Encoding UTF8

        $reader = Import-OfficeCsv -Path $path -AsDataReader -InferSchema -SchemaSampleSize 2
        try {
            $reader.GetType().Name | Should -Be 'CsvDataReader'
            $reader.GetFieldType($reader.GetOrdinal('Value')) | Should -Be ([int])
            $reader.Read() | Should -BeTrue
            $reader.GetString($reader.GetOrdinal('Name')) | Should -Be 'Alpha'
            $reader.GetInt32($reader.GetOrdinal('Value')) | Should -Be 1
        } finally {
            $reader.Dispose()
        }
    }

    It 'imports CSV rows as a typed IDataReader from explicit column types' {
        $path = Join-Path $TestDrive 'reader-column-type.csv'
        Set-Content -LiteralPath $path -Value "Name,Value,Created`nAlpha,1,2026-07-07`nBeta,2,2026-07-08" -Encoding UTF8

        $reader = Import-OfficeCsv -Path $path -AsDataReader -ColumnType @{
            Value = [int]
            Created = [datetime]
        }
        try {
            $reader.GetFieldType($reader.GetOrdinal('Name')) | Should -Be ([string])
            $reader.GetFieldType($reader.GetOrdinal('Value')) | Should -Be ([int])
            $reader.GetFieldType($reader.GetOrdinal('Created')) | Should -Be ([datetime])
            $reader.Read() | Should -BeTrue
            $reader.GetString($reader.GetOrdinal('Name')) | Should -Be 'Alpha'
            $reader.GetInt32($reader.GetOrdinal('Value')) | Should -Be 1
            $reader.GetDateTime($reader.GetOrdinal('Created')) | Should -Be ([datetime]'2026-07-07')
        } finally {
            $reader.Dispose()
        }
    }

    It 'preserves CLR type name casing in explicit column types' {
        $path = Join-Path $TestDrive 'reader-column-type-name.csv'
        Set-Content -LiteralPath $path -Value "Value,Created`n1,2026-07-07" -Encoding UTF8

        $reader = Import-OfficeCsv -Path $path -AsDataReader -ColumnType @{
            Value = 'System.Int32'
            Created = [datetime].AssemblyQualifiedName
        }
        try {
            $reader.GetFieldType($reader.GetOrdinal('Value')) | Should -Be ([int])
            $reader.GetFieldType($reader.GetOrdinal('Created')) | Should -Be ([datetime])
            $reader.Read() | Should -BeTrue
            $reader.GetInt32($reader.GetOrdinal('Value')) | Should -Be 1
            $reader.GetDateTime($reader.GetOrdinal('Created')) | Should -Be ([datetime]'2026-07-07')
        } finally {
            $reader.Dispose()
        }
    }

    It 'rejects combining explicit column types with schema inference' {
        $path = Join-Path $TestDrive 'reader-column-type-infer.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        { Import-OfficeCsv -Path $path -AsDataReader -ColumnType @{ Value = [int] } -InferSchema -ErrorAction Stop } |
            Should -Throw '*ColumnType*InferSchema*'
    }

    It 'keeps progress-enabled CSV readers and documents safe to consume after cmdlet return' {
        $path = Join-Path $TestDrive 'progress-stream.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`nBeta,2" -Encoding UTF8

        $reader = Import-OfficeCsv -Path $path -AsDataReader -Mode Stream -ProgressInterval 1
        try {
            $rows = 0
            while ($reader.Read()) {
                $rows++
            }

            $rows | Should -Be 2
        } finally {
            $reader.Dispose()
        }

        $document = Get-OfficeCsv -Path $path -ProgressInterval 1
        @($document.AsEnumerable()).Count | Should -Be 2
    }

    It 'rejects appending to compressed CSV files' {
        $path = Join-Path $TestDrive 'compressed-append.csv.gz'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType GZip

        {
            [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
                Export-OfficeCsv -Path $path -CompressionType GZip -Append -ErrorAction Stop
        } | Should -Throw '*Appending*compressed*'
    }

    It 'rejects appending to compressed CSV files when compression is inferred from the existing target' {
        $path = Join-Path $TestDrive 'compressed-append-inferred.csv.gz'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType GZip

        {
            [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*Appending*compressed*'
    }

    It 'honors explicit uncompressed appends for extension-based compression names' {
        $path = Join-Path $TestDrive 'plain-extension.csv.gz'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType None

        [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
            Export-OfficeCsv -Path $path -CompressionType None -Append

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'rejects appending to extensionless Deflate compressed CSV files when compression is omitted' {
        $path = Join-Path $TestDrive 'compressed-append-deflate.csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType Deflate

        {
            [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*Appending*compressed*'

        $rows = @(Import-OfficeCsv -Path $path -CompressionType Deflate)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Alpha'
    }

    It 'rejects appending to extensionless Unicode Deflate compressed CSV files when compression is omitted' {
        $path = Join-Path $TestDrive 'compressed-append-deflate-unicode.csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType Deflate -Encoding ([System.Text.Encoding]::Unicode)

        {
            [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*Appending*compressed*'

        $rows = @(Import-OfficeCsv -Path $path -CompressionType Deflate -Encoding ([System.Text.Encoding]::Unicode))
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Alpha'
    }

    It 'rejects appending to extensionless Brotli compressed CSV files when compression is omitted' {
        if (-not $IsCoreCLR) {
            Set-ItResult -Skipped -Because 'Brotli CSV compression is available only on modern .NET runtimes.'
            return
        }

        $path = Join-Path $TestDrive 'compressed-append-brotli.csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType Brotli

        {
            [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*Appending*compressed*'

        $rows = @(Import-OfficeCsv -Path $path -CompressionType Brotli)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Alpha'
    }

    It 'rejects appending to extensionless ZLib compressed CSV files when compression is omitted' {
        if (-not $IsCoreCLR) {
            Set-ItResult -Skipped -Because 'ZLib CSV compression is available only on modern .NET runtimes.'
            return
        }

        $path = Join-Path $TestDrive 'compressed-append-zlib.csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -CompressionType ZLib

        {
            [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*Appending*compressed*'

        $rows = @(Import-OfficeCsv -Path $path -CompressionType ZLib)
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Alpha'
    }

    It 'converts DataTable input directly to CSV text' {
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)

        $csvText = @(ConvertTo-OfficeCsv -InputObject $table)

        $csvText | Should -Be @(
            'Name,Value'
            'Alpha,1'
        )
    }

    It 'converts mixed object and DataTable input without duplicate headers' {
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)

        $csvText = @(
            & {
                Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                Write-Output -InputObject $table -NoEnumerate
            } | ConvertTo-OfficeCsv
        )

        $csvText | Should -Be @(
            'Name,Value'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'converts object then reordered DataTable input using the active column order' {
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Rows.Add(2, 'Beta')

        $csvText = @(
            & {
                Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                Write-Output -InputObject $table -NoEnumerate
            } | ConvertTo-OfficeCsv
        )

        $csvText | Should -Be @(
            'Name,Value'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'converts DataTable then object input using the table column order' {
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)

        $csvText = @(
            & {
                Write-Output -InputObject $table -NoEnumerate
                Write-Output -InputObject ([pscustomobject]@{ Value = 2; Name = 'Beta' }) -NoEnumerate
            } | ConvertTo-OfficeCsv
        )

        $csvText | Should -Be @(
            'Name,Value'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'appends DataTable rows using an existing CSV header order' {
        $path = Join-Path $TestDrive 'datatable-append.csv'
        Set-Content -LiteralPath $path -Value "Value,Name`n1,Alpha" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)

        Export-OfficeCsv -InputObject $table -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @(
            'Value,Name'
            '1,Alpha'
            '2,Beta'
        )
    }

    It 'appends DataTable rows to headerless CSV files without inferring the first data row as a header' {
        $path = Join-Path $TestDrive 'datatable-headerless-append.csv'
        [System.IO.File]::WriteAllText($path, 'Alpha,1')
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)

        Export-OfficeCsv -InputObject $table -Path $path -Append -NoHeader

        Get-Content -LiteralPath $path | Should -Be @(
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'flushes object rows before appending DataTable rows in one invocation' {
        $path = Join-Path $TestDrive 'object-then-datatable-append.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nSeed,0" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)

        & {
            Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
            Write-Output -InputObject $table -NoEnumerate
        } | Export-OfficeCsv -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Seed,0'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'exports object then DataTable rows in one invocation without append' {
        $path = Join-Path $TestDrive 'object-then-datatable-export.csv'
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)

        & {
            Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
            Write-Output -InputObject $table -NoEnumerate
        } | Export-OfficeCsv -Path $path

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'appends DataTable then object rows in one invocation' {
        $path = Join-Path $TestDrive 'datatable-then-object-append.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nSeed,0" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)

        & {
            Write-Output -InputObject $table -NoEnumerate
            Write-Output -InputObject ([pscustomobject]@{ Value = 2; Name = 'Beta' }) -NoEnumerate
        } | Export-OfficeCsv -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Seed,0'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'rejects object then DataTable append when the table is missing active columns' {
        $path = Join-Path $TestDrive 'object-then-datatable-missing-append.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nSeed,0" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Rows.Add('Beta')

        {
            & {
                Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                Write-Output -InputObject $table -NoEnumerate
            } | Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*missing*Value*'

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Seed,0'
            'Alpha,1'
        )
    }

    It 'does not touch an append target when initial DataTable validation fails' {
        $path = Join-Path $TestDrive 'datatable-append-validation-preserve.csv'
        $original = "Name,Value`nAlpha,1"
        [System.IO.File]::WriteAllText($path, $original)
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Rows.Add('Beta')

        {
            Export-OfficeCsv -InputObject $table -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*missing*Value*'

        [System.IO.File]::ReadAllText($path) | Should -Be $original
    }

    It 'uses full matching append headers when DataTable headers are not written' {
        $path = Join-Path $TestDrive 'datatable-append-no-header-full-match.csv'
        Set-Content -LiteralPath $path -Value "Value,Name`n1,Alpha" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)

        Export-OfficeCsv -InputObject $table -Path $path -Append -NoHeader

        Get-Content -LiteralPath $path | Should -Be @(
            'Value,Name'
            '1,Alpha'
            '2,Beta'
        )
    }

    It 'does not infer partial matching headerless DataTable append rows as headers' {
        $path = Join-Path $TestDrive 'datatable-append-no-header-partial-match.csv'
        [System.IO.File]::WriteAllText($path, 'Name,1')
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)

        Export-OfficeCsv -InputObject $table -Path $path -Append -NoHeader

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,1'
            'Beta,2'
        )
    }

    It 'appends IDataReader rows using an existing CSV header order' {
        $path = Join-Path $TestDrive 'reader-append.csv'
        Set-Content -LiteralPath $path -Value "Value,Name`n1,Alpha" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)
        $reader = $table.CreateDataReader()
        try {
            Export-OfficeCsv -InputObject $reader -Path $path -Append
        } finally {
            $reader.Dispose()
        }

        Get-Content -LiteralPath $path | Should -Be @(
            'Value,Name'
            '1,Alpha'
            '2,Beta'
        )
    }

    It 'flushes object rows before appending IDataReader rows in one invocation' {
        $path = Join-Path $TestDrive 'object-then-reader-append.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nSeed,0" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)
        $reader = $table.CreateDataReader()
        try {
            & {
                Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                Write-Output -InputObject $reader -NoEnumerate
            } | Export-OfficeCsv -Path $path -Append
        } finally {
            $reader.Dispose()
        }

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Seed,0'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'preserves headerless append when flushing object rows before IDataReader rows' {
        $path = Join-Path $TestDrive 'object-then-reader-no-header-append.csv'
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)
        $reader = $table.CreateDataReader()
        try {
            & {
                Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                Write-Output -InputObject $reader -NoEnumerate
            } | Export-OfficeCsv -Path $path -Append -NoHeader
        } finally {
            $reader.Dispose()
        }

        Get-Content -LiteralPath $path | Should -Be @(
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'preserves forced headerless append columns when flushing object rows before IDataReader rows' {
        $path = Join-Path $TestDrive 'object-then-reader-no-header-force-append.csv'
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)
        $reader = $table.CreateDataReader()
        try {
            & {
                Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                Write-Output -InputObject $reader -NoEnumerate
            } | Export-OfficeCsv -Path $path -Append -NoHeader -Force
        } finally {
            $reader.Dispose()
        }

        Get-Content -LiteralPath $path | Should -Be @(
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'keeps compressed append streams open when object rows precede IDataReader rows' {
        $path = Join-Path $TestDrive 'object-then-reader-append.csv.gz'
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)
        $reader = $table.CreateDataReader()
        try {
            & {
                Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                Write-Output -InputObject $reader -NoEnumerate
            } | Export-OfficeCsv -Path $path -Append -CompressionType GZip
        } finally {
            $reader.Dispose()
        }

        $rows = @(Import-OfficeCsv -Path $path -CompressionType GZip)
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'Alpha'
        $rows[0].Value | Should -Be '1'
        $rows[1].Name | Should -Be 'Beta'
        $rows[1].Value | Should -Be '2'
    }

    It 'continues appending object rows after IDataReader rows in one invocation' {
        $path = Join-Path $TestDrive 'object-reader-object-append.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nSeed,0" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Beta', 2)
        $reader = $table.CreateDataReader()
        try {
            & {
                Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                Write-Output -InputObject $reader -NoEnumerate
                Write-Output -InputObject ([pscustomobject]@{ Value = 3; Name = 'Gamma' }) -NoEnumerate
            } | Export-OfficeCsv -Path $path -Append
        } finally {
            $reader.Dispose()
        }

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Seed,0'
            'Alpha,1'
            'Beta,2'
            'Gamma,3'
        )
    }

    It 'appends IDataReader then object rows in one invocation' {
        $path = Join-Path $TestDrive 'reader-then-object-append.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nSeed,0" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)
        $reader = $table.CreateDataReader()
        try {
            & {
                Write-Output -InputObject $reader -NoEnumerate
                Write-Output -InputObject ([pscustomobject]@{ Value = 2; Name = 'Beta' }) -NoEnumerate
            } | Export-OfficeCsv -Path $path -Append
        } finally {
            $reader.Dispose()
        }

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Seed,0'
            'Alpha,1'
            'Beta,2'
        )
    }

    It 'rejects object then IDataReader append when the reader is missing active columns' {
        $path = Join-Path $TestDrive 'object-then-reader-missing-append.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nSeed,0" -Encoding UTF8
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Rows.Add('Beta')
        $reader = $table.CreateDataReader()
        try {
            {
                & {
                    Write-Output -InputObject ([pscustomobject]@{ Name = 'Alpha'; Value = 1 }) -NoEnumerate
                    Write-Output -InputObject $reader -NoEnumerate
                } | Export-OfficeCsv -Path $path -Append -ErrorAction Stop
            } | Should -Throw '*missing*Value*'
        } finally {
            $reader.Dispose()
        }

        Get-Content -LiteralPath $path | Should -Be @(
            'Name,Value'
            'Seed,0'
            'Alpha,1'
        )
    }

    It 'does not touch an append target when initial IDataReader validation fails' {
        $path = Join-Path $TestDrive 'reader-append-validation-preserve.csv'
        $original = "Name,Value`nAlpha,1"
        [System.IO.File]::WriteAllText($path, $original)
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Rows.Add('Beta')
        $reader = $table.CreateDataReader()
        try {
            {
                Export-OfficeCsv -InputObject $reader -Path $path -Append -ErrorAction Stop
            } | Should -Throw '*missing*Value*'
        } finally {
            $reader.Dispose()
        }

        [System.IO.File]::ReadAllText($path) | Should -Be $original
    }

    It 'rejects conflicting CSV table and hashtable output modes' {
        $path = Join-Path $TestDrive 'datatable-conflict.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        { Import-OfficeCsv -Path $path -AsDataTable -AsHashtable -ErrorAction Stop } |
            Should -Throw '*AsDataTable*AsDataReader*AsHashtable*'

        { Import-OfficeCsv -Path $path -AsDataTable -AsDataReader -ErrorAction Stop } |
            Should -Throw '*AsDataTable*AsDataReader*AsHashtable*'
    }

    It 'writes to files using the Path alias' {
        $path = Join-Path $TestDrive 'path-alias.csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        Test-Path $path | Should -BeTrue
        Get-Content -Path $path -Raw | Should -Match 'Alpha'
    }

    It 'writes to literal file paths without wildcard expansion' {
        $path = Join-Path $TestDrive 'literal[export].csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -LiteralPath $path

        Test-Path -LiteralPath $path | Should -BeTrue
        (Import-OfficeCsv -LiteralPath $path)[0].Name | Should -Be 'Alpha'
    }

    It 'does not overwrite an existing CSV file when NoClobber is specified' {
        $path = Join-Path $TestDrive 'no-clobber.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nOriginal,1" -Encoding UTF8

        {
            [pscustomobject]@{ Name = 'New'; Value = 2 } |
                Export-OfficeCsv -Path $path -NoClobber -ErrorAction Stop
        } | Should -Throw

        (Import-OfficeCsv -Path $path)[0].Name | Should -Be 'Original'
    }

    It 'appends object rows using the existing CSV header order' {
        $path = Join-Path $TestDrive 'append-order.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        [pscustomobject]@{ Value = 2; Name = 'Beta'; Extra = 'Ignored' } |
            Export-OfficeCsv -Path $path -Append

        $raw = Get-Content -LiteralPath $path
        $data = Import-OfficeCsv -Path $path

        $raw | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
        $data.Count | Should -Be 2
        $data[1].Name | Should -Be 'Beta'
        $data[1].Value | Should -Be '2'
    }

    It 'uses existing CSV header order when appending without writing headers' {
        $path = Join-Path $TestDrive 'append-order-no-header.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        [pscustomobject]@{ Value = 2; Name = 'Beta' } |
            Export-OfficeCsv -Path $path -Append -NoHeader

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
    }

    It 'starts appended rows on a new record when the existing file has no trailing newline' {
        $path = Join-Path $TestDrive 'append-no-newline.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -NoNewline -Encoding UTF8

        [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
            Export-OfficeCsv -Path $path -Append

        $raw = Get-Content -LiteralPath $path -Raw
        $raw | Should -Match "Alpha,1(`r`n|`n|`r)Beta,2"
        $data = @(Import-OfficeCsv -Path $path)
        $data.Count | Should -Be 2
        $data[1].Name | Should -Be 'Beta'
    }

    It 'treats whitespace-only append targets as new CSV files' {
        $path = Join-Path $TestDrive 'append-empty-target.csv'
        Set-Content -LiteralPath $path -Value "`r`n" -NoNewline -Encoding UTF8

        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1')
    }

    It 'does not infer headers when appending to a headerless CSV' {
        $path = Join-Path $TestDrive 'append-headerless.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path -NoHeader

        [pscustomobject]@{ Name = 'Beta'; Value = 2 } |
            Export-OfficeCsv -Path $path -Append -NoHeader

        Get-Content -LiteralPath $path | Should -Be @('Alpha,1', 'Beta,2')
    }

    It 'preserves BOM-detected CSV encoding when appending' {
        $path = Join-Path $TestDrive 'append-utf16.csv'
        $firstName = 'Za' + [char] 0x017C + [char] 0x00F3 + [char] 0x0142 + [char] 0x0107
        $secondName = [string] [char] 0x0141 + [char] 0x00F3 + 'd' + [char] 0x017A
        Set-Content -LiteralPath $path -Value "Name,Value`r`n$firstName,1" -NoNewline -Encoding Unicode

        [pscustomobject]@{ Name = $secondName; Value = 2 } |
            Export-OfficeCsv -Path $path -Append

        $bytes = [System.IO.File]::ReadAllBytes($path)
        $bytes[0] | Should -Be 0xFF
        $bytes[1] | Should -Be 0xFE
        $text = [System.Text.Encoding]::Unicode.GetString($bytes)
        $text | Should -Match ([regex]::Escape("$firstName,1"))
        $text | Should -Match ([regex]::Escape("$secondName,2"))
    }

    It 'appends CLR object rows using existing header casing insensitively' {
        $path = Join-Path $TestDrive 'append-clr-case.csv'
        Set-Content -LiteralPath $path -Value "name,value`nAlpha,1" -Encoding UTF8
        $type = 'PSWriteOffice.Tests.CsvAppendCaseRow' -as [type]
        if (-not $type) {
            $source = "namespace PSWriteOffice.Tests {`n    public sealed class CsvAppendCaseRow {`n        public string Name { get; set; }`n        public int Value { get; set; }`n    }`n}"

            Add-Type -TypeDefinition $source
            $type = 'PSWriteOffice.Tests.CsvAppendCaseRow' -as [type]
        }

        $row = [Activator]::CreateInstance($type)
        $row.Name = 'Beta'
        $row.Value = 2

        $row | Export-OfficeCsv -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @('name,value', 'Alpha,1', 'Beta,2')
    }

    It 'preserves DataRow adapter columns when converting CSV' {
        $table = [System.Data.DataTable]::new()
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Alpha', 1)

        $csvText = @($table.Rows | ConvertTo-OfficeCsv)

        $csvText | Should -Be @('Name,Value', 'Alpha,1')
    }

    It 'requires existing append columns unless Force is specified' {
        $path = Join-Path $TestDrive 'append-force.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        {
            [pscustomobject]@{ Name = 'Beta' } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw

        [pscustomobject]@{ Name = 'Beta' } |
            Export-OfficeCsv -Path $path -Append -Force

        $data = Import-OfficeCsv -Path $path
        $data.Count | Should -Be 2
        $data[1].Name | Should -Be 'Beta'
        $data[1].Value | Should -Be ''
    }

    It 'projects forced scalar appends into an existing Value column' {
        $path = Join-Path $TestDrive 'append-force-scalar.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        'Beta' | Export-OfficeCsv -Path $path -Append -Force

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', ',Beta')
        $data = Import-OfficeCsv -Path $path
        $data[1].Name | Should -Be ''
        $data[1].Value | Should -Be 'Beta'
    }

    It 'projects forced no-header scalar appends into an existing Value column' {
        $path = Join-Path $TestDrive 'append-force-scalar-no-header.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        'Beta' | Export-OfficeCsv -Path $path -Append -NoHeader -Force

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', ',Beta')
        $data = Import-OfficeCsv -Path $path
        $data[1].Name | Should -Be ''
        $data[1].Value | Should -Be 'Beta'
    }

    It 'keeps forced scalar appends on the existing CSV schema without a Value column' {
        $path = Join-Path $TestDrive 'append-force-scalar-schema.csv'
        [pscustomobject]@{ Name = 'Alpha'; Other = 'One' } |
            Export-OfficeCsv -Path $path

        'Beta' | Export-OfficeCsv -Path $path -Append -Force

        Get-Content -LiteralPath $path | Should -Be @('Name,Other', 'Alpha,One', ',')
        $data = @(Import-OfficeCsv -Path $path)
        $data.Count | Should -Be 2
        $data[1].Name | Should -Be ''
        $data[1].Other | Should -Be ''
    }

    It 'keeps mixed scalar conversions on the first row schema without a Value column' {
        $csvText = @(
            [pscustomobject]@{ Name = 'Alpha'; Other = 'One' }
            'Beta'
        ) | ConvertTo-OfficeCsv

        $csvText | Should -Be @('Name,Other', 'Alpha,One', ',')
    }

    It 'does not touch an append target when first row validation fails' {
        $path = Join-Path $TestDrive 'append-validation-preserve.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -NoNewline -Encoding UTF8
        $before = [System.IO.File]::ReadAllBytes($path)

        {
            [pscustomobject]@{ Name = 'Beta' } |
                Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*missing*Value*'

        $after = [System.IO.File]::ReadAllBytes($path)
        [Convert]::ToBase64String($after) | Should -Be ([Convert]::ToBase64String($before))
    }

    It 'validates every appended row against existing columns unless Force is specified' {
        $path = Join-Path $TestDrive 'append-validate-every-row.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        {
            @(
                [pscustomobject]@{ Name = 'Beta'; Value = 2 }
                [pscustomobject]@{ Name = 'Gamma' }
            ) | Export-OfficeCsv -Path $path -Append -ErrorAction Stop
        } | Should -Throw '*missing*Value*'
    }

    It 'appends CSV documents without writing duplicate headers' {
        $path = Join-Path $TestDrive 'append-document.csv'
        [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            Export-OfficeCsv -Path $path

        $document = Get-OfficeCsv -Text "Name,Value`nBeta,2"
        Export-OfficeCsv -Document $document -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
    }

    It 'appends multiple piped CSV documents to the same path' {
        $path = Join-Path $TestDrive 'append-piped-documents.csv'
        $documents = @(
            Get-OfficeCsv -Text "Name,Value`nAlpha,1"
            Get-OfficeCsv -Text "Name,Value`nBeta,2"
        )

        $documents | Export-OfficeCsv -Path $path -Append

        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
    }

    It 'serializes piped CSV documents as rows instead of object properties' {
        $path = Join-Path $TestDrive 'piped-document.csv'
        $document = Get-OfficeCsv -Text "Name,Value`nAlpha,1"

        $csvText = @($document | ConvertTo-OfficeCsv)
        $document | Export-OfficeCsv -Path $path

        $csvText | Should -Be @('Name,Value', 'Alpha,1')
        Get-Content -LiteralPath $path | Should -Be @('Name,Value', 'Alpha,1')
    }

    It 'treats top-level Guid values as scalar CSV values' {
        $guid = [guid]::Parse('00112233-4455-6677-8899-aabbccddeeff')

        $csvText = @($guid | ConvertTo-OfficeCsv -NoHeader)

        $csvText | Should -Be '00112233-4455-6677-8899-aabbccddeeff'
    }

    It 'preserves ETS members added to scalar CSV values' {
        $value = 'Alpha' | Add-Member -NotePropertyName Name -NotePropertyValue 'Decorated' -PassThru

        $csvText = @($value | ConvertTo-OfficeCsv)

        $csvText[0] | Should -BeLike '*Name*'
        $csvText[1] | Should -BeLike '*Decorated*'
    }

    It 'uses the selected culture list separator when UseCulture is specified' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')

        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
            ConvertTo-OfficeCsv -UseCulture -Culture $culture

        $csvText | Should -Contain 'Name;Value'
    }

    It 'uses the selected culture list separator when reading CSV data' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')
        $path = Join-Path $TestDrive 'culture-read.csv'
        Set-Content -LiteralPath $path -Value "Name;Value`nAlpha;1" -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path -UseCulture -Culture $culture)
        $row = $data | Select-Object -First 1

        $data.Count | Should -Be 1
        $row.GetType().FullName | Should -Be 'System.Management.Automation.PSCustomObject'
        $row.Name | Should -Be 'Alpha'
        $row.Value | Should -Be '1'
    }

    It 'detects delimiters for CSV documents and row output' {
        $path = Join-Path $TestDrive 'detect-read.csv'
        Set-Content -LiteralPath $path -Value "Field1;Field2;Field3`n1,2,3,4;5,6,7,8;9,10,11,12" -Encoding UTF8

        $document = Get-OfficeCsv -Path $path -DetectDelimiter
        $data = Import-OfficeCsv -Path $path -DetectDelimiter

        $document.Delimiter | Should -Be ';'
        $document.Header | Should -Be @('Field1', 'Field2', 'Field3')
        $data[0].Field2 | Should -Be '5,6,7,8'
    }

    It 'detects delimiters after skipped preamble rows' {
        $path = Join-Path $TestDrive 'detect-after-preamble.csv'
        Set-Content -LiteralPath $path -Value "generated,by,vendor,with,commas`nName;Value`nAlpha;1" -Encoding UTF8

        $document = Get-OfficeCsv -Path $path -DetectDelimiter -SkipRows 1
        $data = Import-OfficeCsv -Path $path -DetectDelimiter -SkipRows 1

        $document.Delimiter | Should -Be ';'
        $document.Header | Should -Be @('Name', 'Value')
        $data[0].Value | Should -Be '1'
    }

    It 'uses delimiter candidates when detecting from text' {
        $document = Get-OfficeCsv -Text "Name|Value`nAlpha|1" -DetectDelimiter -DelimiterCandidates ';', '|'

        $document.Delimiter | Should -Be '|'
        $row = @($document.AsEnumerable())[0]
        $row['Value'] | Should -Be '1'
    }

    It 'converts CSV text directly into row objects' {
        $data = @(ConvertFrom-OfficeCsv -Text "Name|Value`nAlpha|1" -DetectDelimiter -DelimiterCandidates ';', '|')

        $data.Count | Should -Be 1
        $data[0].Name | Should -Be 'Alpha'
        $data[0].Value | Should -Be '1'
    }

    It 'parses piped CSV text as one stream' {
        $data = "Name,Value", "Alpha,1", "Beta,2" | ConvertFrom-OfficeCsv

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[0].Value | Should -Be '1'
        $data[1].Name | Should -Be 'Beta'
        $data[1].Value | Should -Be '2'
    }

    It 'imports piped file paths as paths rather than CSV text' {
        $path = Join-Path $TestDrive 'piped-path.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        $data = @($path | Import-OfficeCsv)

        $data.Count | Should -Be 1
        $data[0].Name | Should -Be 'Alpha'
    }

    It 'loads CSV documents from literal paths' {
        $path = Join-Path $TestDrive 'literal[1].csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        $document = Get-OfficeCsv -LiteralPath $path

        $document.Header | Should -Be @('Name', 'Value')
        $document.AsEnumerable().Count | Should -Be 1
    }

    It 'loads CSV documents using the legacy InputPath alias' {
        $path = Join-Path $TestDrive 'input-path-alias.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1" -Encoding UTF8

        $document = Get-OfficeCsv -Path $path

        $document.Header | Should -Be @('Name', 'Value')
        $document.AsEnumerable().Count | Should -Be 1
    }

    It 'expands Path wildcards when importing CSV rows' {
        $folder = Join-Path $TestDrive 'wildcard-import'
        New-Item -Path $folder -ItemType Directory | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'a.csv') -Value "Name,Value`nAlpha,1" -Encoding UTF8
        Set-Content -LiteralPath (Join-Path $folder 'b.csv') -Value "Name,Value`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path (Join-Path $folder '*.csv') | Sort-Object Name

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[1].Name | Should -Be 'Beta'
    }

    It 'loads multiple CSV documents from Path values' {
        $folder = Join-Path $TestDrive 'multi-document'
        New-Item -Path $folder -ItemType Directory | Out-Null
        $paths = @(
            Join-Path $folder 'first.csv'
            Join-Path $folder 'second.csv'
        )
        Set-Content -LiteralPath $paths[0] -Value "Name,Value`nAlpha,1" -Encoding UTF8
        Set-Content -LiteralPath $paths[1] -Value "Name,Value`nBeta,2" -Encoding UTF8

        $documents = @(Get-OfficeCsv -Path $paths)

        $documents.Count | Should -Be 2
        $documents[0].Header | Should -Be @('Name', 'Value')
        @($documents[1].AsEnumerable())[0]['Name'] | Should -Be 'Beta'
    }

    It 'preserves unquoted whitespace by default and trims when requested' {
        $path = Join-Path $TestDrive 'whitespace.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,  spaced  " -Encoding UTF8

        $default = Import-OfficeCsv -Path $path
        $trimmed = Import-OfficeCsv -Path $path -TrimWhitespace:$true

        $default[0].Value | Should -Be '  spaced  '
        $trimmed[0].Value | Should -Be 'spaced'
    }

    It 'uses explicit headers and treats the first row as data' {
        $path = Join-Path $TestDrive 'explicit-header.csv'
        Set-Content -LiteralPath $path -Value "Alpha,1`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -Header Name, Value

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[1].Value | Should -Be '2'
    }

    It 'renames duplicate object headers by default and can reject them in strict mode' {
        $path = Join-Path $TestDrive 'duplicate-header.csv'
        Set-Content -LiteralPath $path -Value "Name,Name`nAlpha,1" -Encoding UTF8

        $row = Import-OfficeCsv -Path $path

        $row.Name | Should -Be 'Alpha'
        $row.Name_2 | Should -Be '1'
        { Import-OfficeCsv -Path $path -DuplicateHeaderBehavior Throw -ErrorAction Stop } | Should -Throw '*duplicate*'
        { Import-OfficeCsv -Path $path -DuplicateHeaderBehavior Preserve -ErrorAction Stop } | Should -Throw '*Preserve*row*object*hashtable*'
        { Import-OfficeCsv -Path $path -AsDataTable -DuplicateHeaderBehavior Preserve -ErrorAction Stop } | Should -Throw '*Preserve*DataTable*'
    }

    It 'renames duplicate hashtable headers by default and can reject them in strict mode' {
        $path = Join-Path $TestDrive 'duplicate-hashtable-header.csv'
        Set-Content -LiteralPath $path -Value "Name,Name`nAlpha,1" -Encoding UTF8

        $row = Import-OfficeCsv -Path $path -AsHashtable

        $row['Name'] | Should -Be 'Alpha'
        $row['Name_2'] | Should -Be '1'
        { Import-OfficeCsv -Path $path -AsHashtable -DuplicateHeaderBehavior Throw -ErrorAction Stop } | Should -Throw '*duplicate*'
        { Import-OfficeCsv -Path $path -AsHashtable -DuplicateHeaderBehavior Preserve -ErrorAction Stop } | Should -Throw '*Preserve*row*object*hashtable*'
    }

    It 'rejects duplicate-header preservation for ConvertFrom row outputs' {
        $text = "Name,Name`nAlpha,1"

        { ConvertFrom-OfficeCsv -Text $text -DuplicateHeaderBehavior Preserve -ErrorAction Stop } |
            Should -Throw '*Preserve*row*object*hashtable*'

        { ConvertFrom-OfficeCsv -Text $text -AsHashtable -DuplicateHeaderBehavior Preserve -ErrorAction Stop } |
            Should -Throw '*Preserve*row*object*hashtable*'
    }

    It 'supports NoHeader when reading CSV data and documents' {
        $path = Join-Path $TestDrive 'no-header-read.csv'
        Set-Content -LiteralPath $path -Value "Alpha,1`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -NoHeader
        $document = Get-OfficeCsv -Path $path -NoHeader

        $data.Count | Should -Be 2
        $data[0].Column1 | Should -Be 'Alpha'
        $data[0].Column2 | Should -Be '1'
        $document.Header | Should -Be @('Column1', 'Column2')
        @($document.AsEnumerable()).Count | Should -Be 2
    }

    It 'skips initial records before CSV header discovery' {
        $path = Join-Path $TestDrive 'skip-rows.csv'
        Set-Content -LiteralPath $path -Value "generated by vendor`nexported today`nName,Value`nAlpha,1" -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path -SkipRows 2)
        $document = Get-OfficeCsv -Path $path -SkipRows 2
        $fromText = @(ConvertFrom-OfficeCsv -Text "metadata`nName,Value`nBeta,2" -SkipRows 1)

        $data.Count | Should -Be 1
        $data[0].Name | Should -Be 'Alpha'
        $document.Header | Should -Be @('Name', 'Value')
        $fromText[0].Name | Should -Be 'Beta'
    }

    It 'generates missing header names and tolerates uneven rows by default' {
        $path = Join-Path $TestDrive 'uneven.csv'
        Set-Content -LiteralPath $path -Value "Name,,Value`nAlpha,Ignored`nBeta,Ignored,2,Extra" -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path)
        $rows = @($data | Select-Object -First 2)

        $data.Count | Should -Be 2
        $rows[0].GetType().FullName | Should -Be 'System.Management.Automation.PSCustomObject'
        $rows[0].H1 | Should -Be 'Ignored'
        $rows[0].Value | Should -Be ''
        $rows[1].Value | Should -Be '2'
    }

    It 'keeps trailing missing header properties consistent across import modes' {
        $path = Join-Path $TestDrive 'padded.csv'
        Set-Content -LiteralPath $path -Value "Name,Value,Other`nAlpha,1" -Encoding UTF8

        $streamed = @(Import-OfficeCsv -Path $path)
        $inMemory = @(Import-OfficeCsv -Path $path -Mode InMemory)
        $fromDocument = @(Get-OfficeCsv -Path $path | Import-OfficeCsv)

        foreach ($row in @($streamed[0], $inMemory[0], $fromDocument[0])) {
            $row.PSObject.Properties.Name | Should -Contain 'Other'
            $row.Other | Should -BeNullOrEmpty
        }
    }

    It 'can enforce strict row width validation' {
        $path = Join-Path $TestDrive 'strict-uneven.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha" -Encoding UTF8

        { Import-OfficeCsv -Path $path -ColumnCountMismatchPolicy Strict } | Should -Throw
    }

    It 'recognizes W3C fields headers when loading CSV documents' {
        $document = Get-OfficeCsv -Text "#Version: 1.0`n#Fields: date time cs-uri`n2026-06-24 12:00 /index" -Delimiter ' '

        $document.Header | Should -Be @('date', 'time', 'cs-uri')
        $row = @($document.AsEnumerable())[0]
        $row['cs-uri'] | Should -Be '/index'
    }

    It 'can treat a leading comment row as the header when requested' {
        $document = Get-OfficeCsv -Text "#Name,Value`nAlpha,1" -SkipCommentRowsBeforeHeader:$false

        $document.Header | Should -Be @('#Name', 'Value')
        $row = @($document.AsEnumerable())[0]
        $row['#Name'] | Should -Be 'Alpha'
    }

    It 'does not treat quoted comment-character headers as comments' {
        $path = Join-Path $TestDrive 'quoted-comment-header.csv'
        Set-Content -LiteralPath $path -Value '"#Tag",Name', '10,Alpha' -Encoding UTF8

        $data = @(Import-OfficeCsv -Path $path)
        $document = Get-OfficeCsv -Path $path

        $document.Header | Should -Be @('#Tag', 'Name')
        $data.Count | Should -Be 1
        $data[0].'#Tag' | Should -Be '10'
        $data[0].Name | Should -Be 'Alpha'
    }

    It 'can skip comment rows throughout the file' {
        $path = Join-Path $TestDrive 'comments.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`n# ignored`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -SkipCommentRows

        $data.Count | Should -Be 2
        $data[0].Name | Should -Be 'Alpha'
        $data[1].Name | Should -Be 'Beta'
    }

    It 'can skip custom comment rows throughout the file' {
        $path = Join-Path $TestDrive 'custom-comments.csv'
        Set-Content -LiteralPath $path -Value "Name,Value`nAlpha,1`n; ignored`nBeta,2" -Encoding UTF8

        $data = Import-OfficeCsv -Path $path -SkipCommentRows -CommentCharacter ';'

        $data.Count | Should -Be 2
        $data[1].Name | Should -Be 'Beta'
    }

    It 'lets parameter binding reject Delimiter and UseCulture together' {
        $culture = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')

        {
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 } |
                ConvertTo-OfficeCsv -Delimiter ',' -UseCulture -Culture $culture
        } | Should -Throw
    }

    It 'lets parameter binding reject Delimiter and DetectDelimiter together' {
        {
            Get-OfficeCsv -Text "Name;Value`nAlpha;1" -Delimiter ';' -DetectDelimiter
        } | Should -Throw
    }

    It 'rejects Header and NoHeader together on CSV read surfaces' {
        $path = Join-Path $TestDrive 'header-noheader.csv'
        Set-Content -LiteralPath $path -Value "Alpha,1" -Encoding UTF8

        { Import-OfficeCsv -Path $path -Header Name, Value -NoHeader -ErrorAction Stop } | Should -Throw '*Header*NoHeader*'
        { Get-OfficeCsv -Path $path -Header Name, Value -NoHeader -ErrorAction Stop } | Should -Throw '*Header*NoHeader*'
        { ConvertFrom-OfficeCsv -Text "Alpha,1" -Header Name, Value -NoHeader -ErrorAction Stop } | Should -Throw '*Header*NoHeader*'
    }

    It 'keeps file-only encoding off text parameter sets' {
        $textSets = (Get-Command Get-OfficeCsv).ParameterSets |
            Where-Object Name -like 'Text*'

        foreach ($set in $textSets) {
            $set.Parameters.Name | Should -Not -Contain 'Encoding'
            $set.Parameters.Name | Should -Not -Contain 'CompressionType'
        }

        (Get-Command ConvertTo-OfficeCsv).Parameters.Keys | Should -Not -Contain 'Encoding'
    }

    It 'streams ConvertTo-OfficeCsv output as CSV records that ConvertFrom-OfficeCsv can read' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )

        $csvLines = @($rows | ConvertTo-OfficeCsv)
        $roundTrip = $csvLines | ConvertFrom-OfficeCsv

        $csvLines | Should -Be @('Name,Value', 'Alpha,1', 'Beta,2')
        $roundTrip.Count | Should -Be 2
        $roundTrip[1].Name | Should -Be 'Beta'
        $roundTrip[1].Value | Should -Be '2'
    }

    It 'imports and gets CSV with a multi-character delimiter' {
        $path = Join-Path $TestDrive 'multi-delimiter.csv'
        Set-Content -LiteralPath $path -Value "Name||Value`nAlpha||`"one||two`"" -Encoding UTF8

        $document = Get-OfficeCsv -Path $path -DelimiterText '||'
        $data = @(Import-OfficeCsv -Path $path -DelimiterText '||')

        $document.Header | Should -Be @('Name', 'Value')
        $data.Count | Should -Be 1
        $data[0].Name | Should -Be 'Alpha'
        $data[0].Value | Should -Be 'one||two'
    }

    It 'converts from and to CSV with a multi-character delimiter' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 'one||two' }
            [pscustomobject]@{ Name = 'Beta'; Value = 'plain' }
        )

        $csvLines = @($rows | ConvertTo-OfficeCsv -DelimiterText '||')
        $roundTrip = @($csvLines | ConvertFrom-OfficeCsv -DelimiterText '||')

        $csvLines | Should -Be @('Name||Value', 'Alpha||"one||two"', 'Beta||plain')
        $roundTrip.Count | Should -Be 2
        $roundTrip[0].Value | Should -Be 'one||two'
        $roundTrip[1].Name | Should -Be 'Beta'
    }

    It 'resets multi-character delimiter matching between ConvertTo-OfficeCsv records' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 'ends|' }
            [pscustomobject]@{ Name = ''; Value = "one`ntwo" }
        )

        $csvLines = @($rows | ConvertTo-OfficeCsv -DelimiterText '||')

        $csvLines.Count | Should -Be 3
        $csvLines[1] | Should -Be 'Alpha||ends|'
        $csvLines[2] | Should -Match '^\|\|"one\r?\ntwo"$'
        (@($csvLines | ConvertFrom-OfficeCsv -DelimiterText '||')[1].Value -replace "`r`n", "`n") | Should -Be "one`ntwo"
    }

    It 'exports CSV with a multi-character delimiter and can append using the same delimiter' {
        $path = Join-Path $TestDrive 'export-multi-delimiter.csv'

        [pscustomobject]@{ Name = 'Alpha'; Value = 'one||two' } |
            Export-OfficeCsv -Path $path -DelimiterText '||'
        [pscustomobject]@{ Name = 'Beta'; Value = 'plain' } |
            Export-OfficeCsv -Path $path -DelimiterText '||' -Append

        Get-Content -LiteralPath $path | Should -Be @('Name||Value', 'Alpha||"one||two"', 'Beta||plain')
        @((Import-OfficeCsv -Path $path -DelimiterText '||'))[1].Name | Should -Be 'Beta'
    }

    It 'keeps quoted embedded newlines inside one ConvertTo-OfficeCsv record object' {
        $csvLines = @([pscustomobject]@{ Name = 'Alpha'; Note = "one`ntwo" } | ConvertTo-OfficeCsv)

        $csvLines.Count | Should -Be 2
        $csvLines[0] | Should -Be 'Name,Note'
        $csvLines[1] | Should -Be "Alpha,`"one`ntwo`""
        @($csvLines | ConvertFrom-OfficeCsv)[0].Note | Should -Be "one`ntwo"
    }

    It 'keeps separate records when unquoted values contain quote characters' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Note = 'a"b' }
            [pscustomobject]@{ Name = 'Beta'; Note = 'plain' }
        )

        $csvLines = @($rows | ConvertTo-OfficeCsv -UseQuotes Never)

        $csvLines.Count | Should -Be 3
        $csvLines | Should -Be @('Name,Note', 'Alpha,a"b', 'Beta,plain')
    }

    It 'keeps separate records when unquoted values start with quote characters' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Note = '"starts' }
            [pscustomobject]@{ Name = 'Beta'; Note = 'plain' }
        )

        $csvLines = @($rows | ConvertTo-OfficeCsv -UseQuotes Never)

        $csvLines.Count | Should -Be 3
        $csvLines | Should -Be @('Name,Note', 'Alpha,"starts', 'Beta,plain')
    }

    It 'normalizes CLR projection values and skips failing CLR getters' {
        if (-not ('PSWriteOffice.Tests.CsvClrProjectionRow' -as [type])) {
            Add-Type -TypeDefinition @'
namespace PSWriteOffice.Tests {
    using System;

    public sealed class CsvClrProjectionRow {
        public string Name { get { return "Alpha"; } }
        public string Broken { get { throw new InvalidOperationException("boom"); } }
        public string[] Tags { get { return new[] { "one", "two" }; } }
    }
}
'@
        }

        $row = [PSWriteOffice.Tests.CsvClrProjectionRow]::new()
        $csvLines = @($row | ConvertTo-OfficeCsv)

        $csvLines | Should -Be @('Name,Tags', 'Alpha,"one, two"')
    }

    It 'lets QuoteFields compose with UseQuotes' {
        $path = Join-Path $TestDrive 'quoted-fields.csv'

        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = 1; Note = 'plain' } |
            ConvertTo-OfficeCsv -UseQuotes AsNeeded -QuoteFields Name

        [pscustomobject]@{ Name = 'Alpha'; Value = 1; Note = 'plain' } |
            Export-OfficeCsv -Path $path -UseQuotes AsNeeded -QuoteFields Name

        $csvText | Should -Contain '"Name",Value,Note'
        $csvText | Should -Contain '"Alpha",1,plain'
        (Get-Content -LiteralPath $path -Raw) | Should -Match '"Alpha",1,plain'
    }

    It 'escapes formula-like values when requested' {
        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = '=1+1' } |
            ConvertTo-OfficeCsv -FormulaInjectionPolicy Escape

        ($csvText -join "`n") | Should -Match "'=1\+1"
    }

    It 'uses AsNeeded quoting by default and supports PowerShell-style quote policies' {
        $row = [pscustomobject]@{ Name = 'Alpha'; Value = 'A,B'; Note = 'plain' }

        $default = $row | ConvertTo-OfficeCsv
        $always = $row | ConvertTo-OfficeCsv -UseQuotes Always
        $never = $row | ConvertTo-OfficeCsv -UseQuotes Never
        $quoteFields = $row | ConvertTo-OfficeCsv -QuoteFields Name, Note

        $default | Should -Contain 'Name,Value,Note'
        $default | Should -Contain 'Alpha,"A,B",plain'
        $always | Should -Contain '"Name","Value","Note"'
        $always | Should -Contain '"Alpha","A,B","plain"'
        $never | Should -Contain 'Alpha,A,B,plain'
        $quoteFields | Should -Contain '"Name",Value,"Note"'
        $quoteFields | Should -Contain '"Alpha","A,B","plain"'
    }

    It 'quotes empty values when the quote policy is Always' {
        $csvText = [pscustomobject]@{ Name = 'Alpha'; Value = $null } |
            ConvertTo-OfficeCsv -UseQuotes Always

        $csvText | Should -Contain '"Alpha",""'
    }

    It 'supports NoHeader when converting and exporting CSV' {
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha'; Value = 1 }
            [pscustomobject]@{ Name = 'Beta'; Value = 2 }
        )
        $path = Join-Path $TestDrive 'no-header-export.csv'

        $csvText = $rows | ConvertTo-OfficeCsv -NoHeader
        $rows | Export-OfficeCsv -Path $path -NoHeader

        ($csvText -join "`n") | Should -Not -Match 'Name'
        $csvText | Should -Contain 'Alpha,1'
        (Get-Content -LiteralPath $path -Raw) | Should -Not -Match 'Name'
    }
}
