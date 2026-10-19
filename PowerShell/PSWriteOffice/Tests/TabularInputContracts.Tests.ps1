BeforeAll {
    if (-not $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES) {
        $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
    }
    if (-not $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION) {
        $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = 'Debug'
    }

    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    if (-not ('PSWriteOffice.Tests.TabularContractRow' -as [type])) {
        Add-Type -TypeDefinition @'
using System.Collections.Generic;

namespace PSWriteOffice.Tests {
    public sealed class TabularContractRow {
        public string Name { get; set; }
        public string[] Tags { get; set; }
        public Dictionary<string, object> Metadata { get; set; }
    }

    public sealed class ReadOnlyTabularContractRow : IReadOnlyDictionary<string, object> {
        private readonly KeyValuePair<string, object>[] _entries;
        private readonly Dictionary<string, object> _lookup;

        public ReadOnlyTabularContractRow(IDictionary<string, object> values) {
            _lookup = new Dictionary<string, object>(values, System.StringComparer.OrdinalIgnoreCase);
            _entries = new List<KeyValuePair<string, object>>(values).ToArray();
        }

        public object this[string key] { get { return _lookup[key]; } }
        public IEnumerable<string> Keys { get { foreach (var entry in _entries) yield return entry.Key; } }
        public IEnumerable<object> Values { get { foreach (var entry in _entries) yield return entry.Value; } }
        public int Count { get { return _entries.Length; } }
        public bool ContainsKey(string key) { return _lookup.ContainsKey(key); }
        public bool TryGetValue(string key, out object value) { return _lookup.TryGetValue(key, out value); }
        public IEnumerator<KeyValuePair<string, object>> GetEnumerator() {
            return ((IEnumerable<KeyValuePair<string, object>>)_entries).GetEnumerator();
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { return _entries.GetEnumerator(); }
    }
}
'@
    }

    function New-TestTabularContractReader {
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Tags', [object])
        [void] $table.Columns.Add('Metadata', [object])
        [void] $table.Rows.Add(
            'Alpha',
            [object] @('One', 'Two'),
            [object] ([ordered]@{ Region = 'EU'; Tier = 1 }))
        return ,$table.CreateDataReader()
    }
}

Describe 'Shared tabular input contracts' {
    It 'exposes nested value formatting consistently across table outputs' {
        $commands = @(
            'Add-OfficeExcelTable'
            'Export-OfficeExcel'
            'Add-OfficeExcelReportTable'
            'Add-OfficeWordTable'
            'Add-OfficePowerPointTable'
            'Add-OfficePdfTable'
            'Add-OfficeMarkdownTable'
            'ConvertTo-OfficeMarkdown'
        )

        foreach ($commandName in $commands) {
            $parameters = (Get-Command $commandName -ErrorAction Stop).Parameters.Keys
            $parameters | Should -Contain 'CollectionSeparator'
            $parameters | Should -Contain 'DictionaryEntrySeparator'
            $parameters | Should -Contain 'DictionaryKeyValueSeparator'
        }
    }

    It 'exposes caller-controlled normalization limits across every object table output' {
        $commands = @(
            'Add-OfficeExcelTable'
            'Export-OfficeExcel'
            'Add-OfficeExcelReportTable'
            'Add-OfficeWordTable'
            'Add-OfficePowerPointTable'
            'Add-OfficePdfTable'
            'Add-OfficeMarkdownTable'
            'ConvertTo-OfficeMarkdown'
            'ConvertTo-OfficeCsv'
            'Export-OfficeCsv'
        )

        foreach ($commandName in $commands) {
            $parameters = (Get-Command $commandName -ErrorAction Stop).Parameters.Keys
            $parameters | Should -Contain 'MaxCollectionItems'
            $parameters | Should -Contain 'MaxNestingDepth'
        }

        foreach ($commandName in 'ConvertTo-OfficeCsv', 'Export-OfficeCsv') {
            $documentSets = (Get-Command $commandName -ErrorAction Stop).ParameterSets |
                Where-Object Name -Like 'Document*'
            foreach ($parameterSet in $documentSets) {
                $parameterSet.Parameters.Name | Should -Not -Contain 'MaxCollectionItems'
                $parameterSet.Parameters.Name | Should -Not -Contain 'MaxNestingDepth'
            }
        }
    }

    It 'reports and honors the required collection limit on Excel report tables' {
        $row = [pscustomobject]@{ Name = 'Alpha'; Tags = @('One', 'Two', 'Three') }
        $limitedPath = Join-Path $TestDrive 'Report-CollectionLimit.xlsx'

        {
            New-OfficeExcel -Path $limitedPath {
                Add-OfficeExcelReportSheet -Name 'Report' {
                    Add-OfficeExcelReportTable -InputObject $row -MaxCollectionItems 2
                }
            }
        } | Should -Throw '*collection*at least 3 items*-MaxCollectionItems 2*Rerun with -MaxCollectionItems 3 or higher*'

        New-OfficeExcel -Path $limitedPath {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $row -MaxCollectionItems 3
            }
        }

        $result = @(Import-OfficeExcel -Path $limitedPath -WorksheetName 'Report')
        $result[0].Tags | Should -Be 'One, Two, Three'
    }

    It 'counts container nesting consistently regardless of the scalar leaf type' {
        $leafValues = @('Deep', 42, $true, [datetime]'2026-08-12T00:00:00Z')
        foreach ($leafValue in $leafValues) {
            $row = [pscustomobject]@{
                Name   = 'Alpha'
                Nested = [ordered]@{ Child = [ordered]@{ Value = $leafValue } }
            }

            {
                $row | ConvertTo-OfficeMarkdown -MaxNestingDepth 1 -ErrorAction Stop
            } | Should -Throw '*requires at least 2 normalization levels*-MaxNestingDepth 1*Rerun with -MaxNestingDepth 2 or higher*'

            { $row | ConvertTo-OfficeMarkdown -MaxNestingDepth 2 -ErrorAction Stop } |
                Should -Not -Throw
        }
    }

    It 'routes collection limits through direct Excel exports' {
        $row = [pscustomobject]@{ Name = 'Alpha'; Tags = @('One', 'Two', 'Three') }
        $path = Join-Path $TestDrive 'Export-CollectionLimit.xlsx'

        {
            $row | Export-OfficeExcel -Path $path -MaxCollectionItems 2 -ErrorAction Stop
        } | Should -Throw '*collection*at least 3 items*-MaxCollectionItems 2*Rerun with -MaxCollectionItems 3 or higher*'

        $row | Export-OfficeExcel -Path $path -MaxCollectionItems 3
        $result = @(Import-OfficeExcel -Path $path -WorksheetName 'Sheet1')
        $result[0].Tags | Should -Be 'One, Two, Three'
    }

    It 'applies normalization limits to every DataSet table' {
        $dataSet = [System.Data.DataSet]::new('Book')
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Tags', [object])
        [void] $table.Rows.Add('Alpha', [object] @('One', 'Two', 'Three'))
        [void] $dataSet.Tables.Add($table)
        $path = Join-Path $TestDrive 'DataSet-CollectionLimit.xlsx'

        {
            $dataSet | Export-OfficeExcel -Path $path -MaxCollectionItems 2 -ErrorAction Stop
        } | Should -Throw '*collection*at least 3 items*-MaxCollectionItems 2*Rerun with -MaxCollectionItems 3 or higher*'

        $dataSet | Export-OfficeExcel -Path $path -MaxCollectionItems 3
        $result = @(Import-OfficeExcel -Path $path -WorksheetName 'Rows')
        $result[0].Tags | Should -Be 'One, Two, Three'
    }

    It 'does not apply nested-cell collection limits to top-level dictionary columns' {
        $row = [ordered]@{ A = 1; B = 2; C = 3 }
        $path = Join-Path $TestDrive 'TopLevelDictionary.xlsx'

        { $row | Export-OfficeExcel -Path $path -MaxCollectionItems 2 -ErrorAction Stop } |
            Should -Not -Throw

        $result = @(Import-OfficeExcel -Path $path -WorksheetName 'Sheet1')
        $result.Count | Should -Be 1
        $result[0].A | Should -Be 1
        $result[0].B | Should -Be 2
        $result[0].C | Should -Be 3
    }

    It 'routes dictionary limits through streaming CSV exports' {
        $row = [pscustomobject]@{
            Name     = 'Alpha'
            Metadata = [ordered]@{ First = 1; Second = 2; Third = 3 }
        }

        {
            $row | ConvertTo-OfficeCsv -MaxCollectionItems 2 -ErrorAction Stop
        } | Should -Throw '*dictionary*at least 3 items*-MaxCollectionItems 2*Rerun with -MaxCollectionItems 3 or higher*'

        $csv = @($row | ConvertTo-OfficeCsv -MaxCollectionItems 3)
        ($csv -join "`n") | Should -Match 'First: 1; Second: 2; Third: 3'
    }

    It 'reuses raised nesting limits while validating CSV append columns' {
        $path = Join-Path $TestDrive 'Append-RaisedNestingLimit.csv'
        [System.IO.File]::WriteAllText($path, "Name,Payload`r`nAlpha,Initial`r`n")

        $payload = 'Deep'
        foreach ($level in 1..70) {
            $payload = [ordered]@{ Child = $payload }
        }
        $row = [pscustomobject]@{ Name = 'Beta'; Payload = $payload }

        { $row | Export-OfficeCsv -Path $path -Append -MaxNestingDepth 128 -ErrorAction Stop } |
            Should -Not -Throw

        $result = @(Import-OfficeCsv -Path $path)
        $result.Count | Should -Be 2
        $result[1].Name | Should -Be 'Beta'
        $result[1].Payload | Should -Match 'Deep'
    }

    It 'routes collection limits through CSV DataTable and DataView fast paths' {
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Tags', [object])
        [void] $table.Rows.Add('Alpha', [object] @('One', 'Two', 'Three'))

        {
            ConvertTo-OfficeCsv -InputObject $table -MaxCollectionItems 2 -ErrorAction Stop
        } | Should -Throw '*collection*at least 3 items*-MaxCollectionItems 2*Rerun with -MaxCollectionItems 3 or higher*'

        $csv = @(ConvertTo-OfficeCsv -InputObject $table -MaxCollectionItems 3)
        ($csv -join "`n") | Should -Match 'One, Two, Three'

        $path = Join-Path $TestDrive 'DataView-CollectionLimit.csv'
        {
            Export-OfficeCsv -InputObject $table.DefaultView -Path $path -MaxCollectionItems 2 -ErrorAction Stop
        } | Should -Throw '*collection*at least 3 items*-MaxCollectionItems 2*Rerun with -MaxCollectionItems 3 or higher*'

        Export-OfficeCsv -InputObject $table.DefaultView -Path $path -MaxCollectionItems 3
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'One, Two, Three'
    }

    It 'routes dictionary limits through the CSV IDataReader fast path' {
        $table = [System.Data.DataTable]::new('Rows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Metadata', [object])
        [void] $table.Rows.Add('Alpha', [object] ([ordered]@{ First = 1; Second = 2; Third = 3 }))
        $path = Join-Path $TestDrive 'DataReader-CollectionLimit.csv'

        $limitedReader = $table.CreateDataReader()
        try {
            {
                Export-OfficeCsv -InputObject $limitedReader -Path $path -MaxCollectionItems 2 -ErrorAction Stop
            } | Should -Throw '*dictionary*at least 3 items*-MaxCollectionItems 2*Rerun with -MaxCollectionItems 3 or higher*'
        } finally {
            $limitedReader.Dispose()
        }

        $allowedReader = $table.CreateDataReader()
        try {
            Export-OfficeCsv -InputObject $allowedReader -Path $path -MaxCollectionItems 3
        } finally {
            $allowedReader.Dispose()
        }
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'First: 1; Second: 2; Third: 3'
    }

    It 'renders object and dictionary row families through Excel report tables' {
        $genericDictionary = [System.Collections.Generic.Dictionary[string, object]]::new()
        $genericDictionary.Add('Name', 'GenericDictionary')
        $genericDictionary.Add('Value', 3)

        $cases = @(
            [pscustomobject]@{
                Name  = 'PSCustomObject'
                Input = [pscustomobject]@{ Name = 'PSCustomObject'; Value = 1 }
            }
            [pscustomobject]@{
                Name  = 'OrderedDictionary'
                Input = [ordered]@{ Name = 'OrderedDictionary'; Value = 2 }
            }
            [pscustomobject]@{
                Name  = 'GenericDictionary'
                Input = $genericDictionary
            }
            [pscustomobject]@{
                Name  = 'ClrObject'
                Input = [System.Collections.DictionaryEntry]::new('ClrObject', 4)
            }
        )

        foreach ($case in $cases) {
            $path = Join-Path $TestDrive ("Report-{0}.xlsx" -f $case.Name)
            $inputRow = $case.Input

            New-OfficeExcel -Path $path {
                Add-OfficeExcelReportSheet -Name 'Report' {
                    Add-OfficeExcelReportTable -InputObject $inputRow
                }
            }

            $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
            $rows.Count | Should -Be 1

            if ($case.Name -eq 'ClrObject') {
                $rows[0].Key | Should -Be 'ClrObject'
                $rows[0].Value | Should -Be 4
            } else {
                $rows[0].Name | Should -Be $case.Name
                $rows[0].Value | Should -BeGreaterThan 0
            }
        }
    }

    It 'unions columns found only in later heterogeneous PowerShell rows' {
        $path = Join-Path $TestDrive 'Report-HeterogeneousRows.xlsx'
        $rows = @(
            [pscustomobject]@{ Name = 'Alpha' }
            [pscustomobject]@{ Name = 'Beta'; Department = 'Operations' }
            [ordered]@{ Name = 'Gamma'; Score = 7 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $rows
            }
        }

        $result = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
        $result | Should -HaveCount 3
        $result[0].PSObject.Properties.Name | Should -Contain 'Department'
        $result[0].PSObject.Properties.Name | Should -Contain 'Score'
        $result[0].Department | Should -BeNullOrEmpty
        $result[1].Department | Should -Be 'Operations'
        $result[2].Score | Should -Be 7
    }

    It 'renders ADO.NET tabular input families without DataRow metadata' {
        $createTable = {
            $table = [System.Data.DataTable]::new('Rows')
            [void] $table.Columns.Add('Name', [string])
            [void] $table.Columns.Add('Value', [int])
            [void] $table.Rows.Add('Alpha', 7)
            return ,$table
        }

        $inputFactories = [ordered]@{
            DataTable   = { & $createTable }
            DataView    = { (& $createTable).DefaultView }
            DataRow     = { (& $createTable).Rows[0] }
            IDataReader = { return ,((& $createTable).CreateDataReader()) }
        }

        foreach ($entry in $inputFactories.GetEnumerator()) {
            $path = Join-Path $TestDrive ("Report-{0}.xlsx" -f $entry.Key)
            $inputRows = & $entry.Value
            try {
                New-OfficeExcel -Path $path {
                    Add-OfficeExcelReportSheet -Name 'Report' {
                        Add-OfficeExcelReportTable -InputObject $inputRows
                    }
                }
            } finally {
                if ($inputRows -is [System.Data.IDataReader]) {
                    $inputRows.Dispose()
                }
            }

            $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
            $rows.Count | Should -Be 1 -Because $entry.Key
            $rows[0].Name | Should -Be 'Alpha' -Because $entry.Key
            $rows[0].Value | Should -Be 7 -Because $entry.Key
            $rows[0].PSObject.Properties.Name | Should -Not -Contain 'RowError'
            $rows[0].PSObject.Properties.Name | Should -Not -Contain 'Table'
        }
    }

    It 'renders scalar input as a single Value column' {
        $path = Join-Path $TestDrive 'Report-Scalar.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject 'Scalar value'
            }
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
        $rows.Count | Should -Be 1
        $rows[0].Value | Should -Be 'Scalar value'
    }

    It 'formats nested collections and dictionaries identically in Excel and Markdown tables' {
        $inputRow = [pscustomobject]@{
            Name     = 'Alpha'
            Tags     = @('One', 'Two')
            Metadata = [ordered]@{ Region = 'EU'; Tier = 1 }
        }
        $excelPath = Join-Path $TestDrive 'Report-NestedValues.xlsx'

        New-OfficeExcel -Path $excelPath {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $inputRow `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            }
        }

        $excelRows = @(Import-OfficeExcel -Path $excelPath -WorksheetName 'Report')
        $excelRows[0].Tags | Should -Be 'One | Two'
        $excelRows[0].Metadata | Should -Be 'Region: EU; Tier: 1'

        $markdownPath = Join-Path $TestDrive 'Report-NestedValues.md'
        New-OfficeMarkdown -Path $markdownPath {
            Add-OfficeMarkdownTable -InputObject $inputRow `
                -CollectionSeparator ' | ' `
                -DictionaryEntrySeparator '; ' `
                -DictionaryKeyValueSeparator ': '
        }

        $markdown = Get-Content -LiteralPath $markdownPath -Raw
        $markdown | Should -Match ([regex]::Escape('One \| Two'))
        $markdown | Should -Match 'Region: EU; Tier: 1'

        $convertedMarkdown = $inputRow | ConvertTo-OfficeMarkdown `
            -CollectionSeparator ' | ' `
            -DictionaryEntrySeparator '; ' `
            -DictionaryKeyValueSeparator ': '
        $convertedMarkdown | Should -Match ([regex]::Escape('One \| Two'))
        $convertedMarkdown | Should -Match 'Region: EU; Tier: 1'
    }

    It 'supports column inclusion and exclusion on Excel report tables' {
        $path = Join-Path $TestDrive 'Report-Projection.xlsx'
        $inputRow = [pscustomobject]@{ Name = 'Alpha'; Value = 1; Internal = 'Hidden' }

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $inputRow -Property Value, Name -ExcludeProperty Internal
            }
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
        $rows[0].PSObject.Properties.Name | Should -Be @('Value', 'Name')
        $rows[0].Value | Should -Be 1
        $rows[0].Name | Should -Be 'Alpha'
    }

    It 'preserves incoming column order on Excel report tables by default' {
        $path = Join-Path $TestDrive 'Report-SourceOrder.xlsx'
        $inputRow = [pscustomobject] [ordered] @{
            Zulu   = 'First'
            Alpha  = 'Second'
            Middle = 'Third'
        }

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $inputRow
            }
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Report')
        $rows[0].PSObject.Properties.Name | Should -Be @('Zulu', 'Alpha', 'Middle')
    }

    It 'streams IDataReader rows with shared nested formatting into direct Excel exports' {
        $path = Join-Path $TestDrive 'DirectReader.xlsx'
        $reader = New-TestTabularContractReader
        try {
            Export-OfficeExcel -Path $path -InputObject $reader `
                -CollectionSeparator ' | ' `
                -DictionaryEntrySeparator '; ' `
                -DictionaryKeyValueSeparator ': '
        } finally {
            $reader.Dispose()
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sheet1')
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Alpha'
        $rows[0].Tags | Should -Be 'One | Two'
        $rows[0].Metadata | Should -Be 'Region: EU; Tier: 1'
    }

    It 'normalizes dictionary fast-path exports before OfficeIMO writes the workbook' {
        $path = Join-Path $TestDrive 'DirectDictionary.xlsx'
        $row = [ordered]@{
            Name = 'Alpha'
            Tags = @('One', 'Two')
            Metadata = [ordered]@{ Region = 'EU'; Tier = 1 }
        }

        Export-OfficeExcel -Path $path -InputObject $row `
            -CollectionSeparator ' | ' `
            -DictionaryEntrySeparator '; ' `
            -DictionaryKeyValueSeparator ': '

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sheet1')
        $rows[0].Tags | Should -Be 'One | Two'
        $rows[0].Metadata | Should -Be 'Region: EU; Tier: 1'
    }

    It 'materializes IDataReader rows consistently across Word PowerPoint PDF and Markdown tables' {
        $wordPath = Join-Path $TestDrive 'Reader.docx'
        $reader = New-TestTabularContractReader
        try {
            New-OfficeWord -Path $wordPath {
                Add-OfficeWordTable -InputObject $reader -Style TableGrid `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            } | Out-Null
        } finally {
            $reader.Dispose()
        }

        $word = Get-OfficeWord -Path $wordPath -ReadOnly
        try {
            $word.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Alpha'
            $word.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'One | Two'
            $word.Tables[0].Rows[1].Cells[2].Paragraphs[0].Text | Should -Be 'Region: EU; Tier: 1'
        } finally {
            $word.Dispose()
        }

        $powerPointPath = Join-Path $TestDrive 'Reader.pptx'
        $presentation = New-OfficePowerPoint -Path $powerPointPath -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -PassThru
            $reader = New-TestTabularContractReader
            try {
                $table = Add-OfficePowerPointTable -Slide $slide -InputObject $reader `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': ' -PassThru
            } finally {
                $reader.Dispose()
            }

            $table.GetCell(1, 0).Text | Should -Be 'Alpha'
            $table.GetCell(1, 1).Text | Should -Be 'One | Two'
            $table.GetCell(1, 2).Text | Should -Be 'Region: EU; Tier: 1'
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }

        $pdfPath = Join-Path $TestDrive 'Reader.pdf'
        $reader = New-TestTabularContractReader
        try {
            New-OfficePdf -Path $pdfPath {
                Add-OfficePdfTable -InputObject $reader `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            } | Out-Null
        } finally {
            $reader.Dispose()
        }

        $pdfText = Get-OfficePdfText -Path $pdfPath
        $pdfText | Should -Match 'Alpha'
        $pdfText | Should -Match 'One \| Two'
        $pdfText | Should -Match 'Region: EU; Tier: 1'

        $markdownPath = Join-Path $TestDrive 'Reader.md'
        $reader = New-TestTabularContractReader
        try {
            New-OfficeMarkdown -Path $markdownPath {
                Add-OfficeMarkdownTable -InputObject $reader `
                    -CollectionSeparator ' | ' `
                    -DictionaryEntrySeparator '; ' `
                    -DictionaryKeyValueSeparator ': '
            }
        } finally {
            $reader.Dispose()
        }

        $markdown = Get-Content -LiteralPath $markdownPath -Raw
        $markdown | Should -Match 'Alpha'
        $markdown | Should -Match ([regex]::Escape('One \| Two'))
        $markdown | Should -Match 'Region: EU; Tier: 1'
    }

    It 'projects CLR rows before nested values reach format-specific writers' {
        $metadata = [System.Collections.Generic.Dictionary[string, object]]::new()
        $metadata.Add('Region', 'EU')
        $metadata.Add('Tier', 1)
        $row = [PSWriteOffice.Tests.TabularContractRow]::new()
        $row.Name = 'Alpha'
        $row.Tags = @('One', 'Two')
        $row.Metadata = $metadata

        $markdown = $row | ConvertTo-OfficeMarkdown `
            -CollectionSeparator ' | ' `
            -DictionaryEntrySeparator '; ' `
            -DictionaryKeyValueSeparator ': '

        $markdown | Should -Match 'Alpha'
        $markdown | Should -Match ([regex]::Escape('One \| Two'))
        $markdown | Should -Match 'Region: EU; Tier: 1'
    }

    It 'keeps IReadOnlyDictionary-only rows as named columns across table formats' {
        $values = [System.Collections.Generic.Dictionary[string, object]]::new()
        $values.Add('Name', 'Alpha')
        $values.Add('Tags', @('One', 'Two'))
        $values.Add('Metadata', [ordered]@{ Region = 'EU'; Tier = 1 })
        $row = [PSWriteOffice.Tests.ReadOnlyTabularContractRow]::new($values)

        $exportPath = Join-Path $TestDrive 'ReadOnlyDictionaryExport.xlsx'
        Export-OfficeExcel -Path $exportPath -InputObject $row -CollectionSeparator ' | '
        $exportRows = @(Import-OfficeExcel -Path $exportPath -WorksheetName 'Sheet1')
        $exportRows[0].Name | Should -Be 'Alpha'
        $exportRows[0].Tags | Should -Be 'One | Two'

        $tablePath = Join-Path $TestDrive 'ReadOnlyDictionaryTable.xlsx'
        New-OfficeExcel -Path $tablePath {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $row -TableName 'ReadOnlyRows' -CollectionSeparator ' | '
            }
        }
        $tableRows = @(Import-OfficeExcel -Path $tablePath -WorksheetName 'Data')
        $tableRows[0].Name | Should -Be 'Alpha'
        $tableRows[0].Tags | Should -Be 'One | Two'

        $reportPath = Join-Path $TestDrive 'ReadOnlyDictionaryReport.xlsx'
        New-OfficeExcel -Path $reportPath {
            Add-OfficeExcelReportSheet -Name 'Report' {
                Add-OfficeExcelReportTable -InputObject $row -CollectionSeparator ' | '
            }
        }
        $reportRows = @(Import-OfficeExcel -Path $reportPath -WorksheetName 'Report')
        $reportRows[0].Name | Should -Be 'Alpha'
        $reportRows[0].Tags | Should -Be 'One | Two'

        $markdown = $row | ConvertTo-OfficeMarkdown -CollectionSeparator ' | '
        $markdown | Should -Match 'Name'
        $markdown | Should -Match 'Alpha'
        $markdown | Should -Match ([regex]::Escape('One \| Two'))

        $wordPath = Join-Path $TestDrive 'ReadOnlyDictionary.docx'
        New-OfficeWord -Path $wordPath {
            Add-OfficeWordTable -InputObject $row -Style TableGrid -CollectionSeparator ' | '
        } | Out-Null
        $word = Get-OfficeWord -Path $wordPath -ReadOnly
        try {
            $word.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text | Should -Be 'Name'
            $word.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text | Should -Be 'Alpha'
            $word.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text | Should -Be 'One | Two'
        } finally {
            $word.Dispose()
        }

        $powerPointPath = Join-Path $TestDrive 'ReadOnlyDictionary.pptx'
        $presentation = New-OfficePowerPoint -Path $powerPointPath -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -PassThru
            $table = Add-OfficePowerPointTable -Slide $slide -InputObject $row -CollectionSeparator ' | ' -PassThru
            $table.GetCell(0, 0).Text | Should -Be 'Name'
            $table.GetCell(1, 0).Text | Should -Be 'Alpha'
            $table.GetCell(1, 1).Text | Should -Be 'One | Two'
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }

        $pdfPath = Join-Path $TestDrive 'ReadOnlyDictionary.pdf'
        New-OfficePdf -Path $pdfPath {
            Add-OfficePdfTable -InputObject $row -CollectionSeparator ' | '
        } | Out-Null
        $pdfText = Get-OfficePdfText -Path $pdfPath
        $pdfText | Should -Match 'Name'
        $pdfText | Should -Match 'Alpha'
        $pdfText | Should -Match 'One \| Two'
    }
}
