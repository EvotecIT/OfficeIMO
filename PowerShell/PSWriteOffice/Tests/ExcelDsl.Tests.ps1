BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')

    if (-not ('PSWriteOfficeTestSupport.LateDestinationDataReader' -as [type])) {
        $readerTypeDefinition = @'
using System;
using System.Data;
using System.IO;

namespace PSWriteOfficeTestSupport {
    public sealed class LateDestinationDataReader : IDataReader {
        private readonly IDataReader _inner;
        private readonly string _path;
        private readonly string _content;
        private bool _created;

        public LateDestinationDataReader(IDataReader inner, string path, string content) {
            _inner = inner;
            _path = path;
            _content = content;
        }

        public object this[int i] { get { return _inner[i]; } }
        public object this[string name] { get { return _inner[name]; } }
        public int Depth { get { return _inner.Depth; } }
        public bool IsClosed { get { return _inner.IsClosed; } }
        public int RecordsAffected { get { return _inner.RecordsAffected; } }
        public int FieldCount { get { return _inner.FieldCount; } }
        public void Close() { _inner.Close(); }
        public void Dispose() { _inner.Dispose(); }
        public bool GetBoolean(int i) { return _inner.GetBoolean(i); }
        public byte GetByte(int i) { return _inner.GetByte(i); }
        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferOffset, int length) { return _inner.GetBytes(i, fieldOffset, buffer, bufferOffset, length); }
        public char GetChar(int i) { return _inner.GetChar(i); }
        public long GetChars(int i, long fieldOffset, char[] buffer, int bufferOffset, int length) { return _inner.GetChars(i, fieldOffset, buffer, bufferOffset, length); }
        public IDataReader GetData(int i) { return _inner.GetData(i); }
        public string GetDataTypeName(int i) { return _inner.GetDataTypeName(i); }
        public DateTime GetDateTime(int i) { return _inner.GetDateTime(i); }
        public decimal GetDecimal(int i) { return _inner.GetDecimal(i); }
        public double GetDouble(int i) { return _inner.GetDouble(i); }
        public Type GetFieldType(int i) { return _inner.GetFieldType(i); }
        public float GetFloat(int i) { return _inner.GetFloat(i); }
        public Guid GetGuid(int i) { return _inner.GetGuid(i); }
        public short GetInt16(int i) { return _inner.GetInt16(i); }
        public int GetInt32(int i) { return _inner.GetInt32(i); }
        public long GetInt64(int i) { return _inner.GetInt64(i); }
        public string GetName(int i) { return _inner.GetName(i); }
        public int GetOrdinal(string name) { return _inner.GetOrdinal(name); }
        public DataTable GetSchemaTable() { return _inner.GetSchemaTable(); }
        public string GetString(int i) { return _inner.GetString(i); }
        public object GetValue(int i) { return _inner.GetValue(i); }
        public int GetValues(object[] values) { return _inner.GetValues(values); }
        public bool IsDBNull(int i) { return _inner.IsDBNull(i); }
        public bool NextResult() { return _inner.NextResult(); }

        public bool Read() {
            if (!_created) {
                File.WriteAllText(_path, _content);
                _created = true;
            }

            return _inner.Read();
        }
    }
}
'@
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            Add-Type -TypeDefinition $readerTypeDefinition -ReferencedAssemblies @(
                'System.Data.dll'
                'System.Xml.dll'
            )
        } else {
            Add-Type -TypeDefinition $readerTypeDefinition
        }
    }

    function Get-TestLoadedType {
        param(
            [Parameter(Mandatory)]
            [string] $Name
        )

        $type = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType($Name, $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        if ($null -eq $type) {
            throw "Unable to find loaded type '$Name'."
        }

        $type
    }

    function Test-OfficeLoadedMethod {
        param(
            [Parameter(Mandatory)]
            [string] $TypeName,

            [Parameter(Mandatory)]
            [string] $MethodName
        )

        $type = Get-TestLoadedType -Name $TypeName
        @($type.GetMethods() | Where-Object Name -eq $MethodName).Count -gt 0
    }

    function Get-TestWorkbookDate1904 {
        param(
            [Parameter(Mandatory)]
            [string] $Path
        )

        $workbook = Get-ZipXmlDocumentLocal -Path $Path -Entry 'xl/workbook.xml'
        $workbook.GetElementsByTagName('workbookPr', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')[0].date1904
    }

    function Get-TestWorksheetCellValue {
        param(
            [Parameter(Mandatory)]
            [string] $Path,

            [Parameter(Mandatory)]
            [string] $Address,

            [string] $Entry = 'xl/worksheets/sheet1.xml'
        )

        $worksheet = Get-ZipXmlDocumentLocal -Path $Path -Entry $Entry
        $cell = $worksheet.GetElementsByTagName('c', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') |
            Where-Object { $_.r -eq $Address } |
            Select-Object -First 1
        $cell.GetElementsByTagName('v', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')[0].'#text'
    }

    function Get-TestWorksheetCellStyleIndex {
        param(
            [Parameter(Mandatory)]
            [string] $Path,

            [Parameter(Mandatory)]
            [string] $Address,

            [string] $Entry = 'xl/worksheets/sheet1.xml'
        )

        $worksheet = Get-ZipXmlDocumentLocal -Path $Path -Entry $Entry
        $cell = $worksheet.GetElementsByTagName('c', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') |
            Where-Object { $_.r -eq $Address } |
            Select-Object -First 1
        $cell.s
    }

    function Get-TestWorksheetCellFormula {
        param(
            [Parameter(Mandatory)]
            [string] $Path,

            [Parameter(Mandatory)]
            [string] $Address,

            [string] $Entry = 'xl/worksheets/sheet1.xml'
        )

        $worksheet = Get-ZipXmlDocumentLocal -Path $Path -Entry $Entry
        $cell = $worksheet.GetElementsByTagName('c', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') |
            Where-Object { $_.r -eq $Address } |
            Select-Object -First 1
        $cell.GetElementsByTagName('f', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')[0].'#text'
    }

    function Get-TestWorksheetPageSetupAttribute {
        param(
            [Parameter(Mandatory)]
            [string] $Path,

            [Parameter(Mandatory)]
            [string] $Name,

            [string] $Entry = 'xl/worksheets/sheet1.xml'
        )

        $worksheet = Get-ZipXmlDocumentLocal -Path $Path -Entry $Entry
        $pageSetup = $worksheet.GetElementsByTagName('pageSetup', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') |
            Select-Object -First 1
        $pageSetup.$Name
    }

    function Get-TestWorksheetPageSetupPropertyAttribute {
        param(
            [Parameter(Mandatory)]
            [string] $Path,

            [Parameter(Mandatory)]
            [string] $Name,

            [string] $Entry = 'xl/worksheets/sheet1.xml'
        )

        $worksheet = Get-ZipXmlDocumentLocal -Path $Path -Entry $Entry
        $pageSetupProperties = $worksheet.GetElementsByTagName('pageSetUpPr', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') |
            Select-Object -First 1
        $pageSetupProperties.$Name
    }

    function Get-TestWorksheetPageMarginAttribute {
        param(
            [Parameter(Mandatory)]
            [string] $Path,

            [Parameter(Mandatory)]
            [string] $Name,

            [string] $Entry = 'xl/worksheets/sheet1.xml'
        )

        $worksheet = Get-ZipXmlDocumentLocal -Path $Path -Entry $Entry
        $pageMargins = $worksheet.GetElementsByTagName('pageMargins', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') |
            Select-Object -First 1
        $pageMargins.$Name
    }

    function Get-TestWorkbookDefinedName {
        param(
            [Parameter(Mandatory)]
            [string] $Path,

            [Parameter(Mandatory)]
            [string] $Name
        )

        $workbook = Get-ZipXmlDocumentLocal -Path $Path -Entry 'xl/workbook.xml'
        $workbook.GetElementsByTagName('definedName', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main') |
            Where-Object { $_.name -eq $Name } |
            Select-Object -First 1 |
            ForEach-Object { $_.'#text' }
    }

    function Get-TestExcel1904Serial {
        param(
            [Parameter(Mandatory)]
            [datetime] $Value
        )

        $Value.ToOADate() - 1462
    }

    function Add-TestWorkbookInteractionParts {
        param(
            [Parameter(Mandatory)]
            [string] $Path
        )

        if (-not ('DocumentFormat.OpenXml.Packaging.SpreadsheetDocument' -as [type])) {
            $configuration = if ($env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION -in @('Debug', 'Release')) {
                $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION
            } elseif ($env:BUILD_CONFIGURATION -in @('Debug', 'Release')) {
                $env:BUILD_CONFIGURATION
            } elseif (Test-Path (Join-Path $PSScriptRoot '..\Sources\PSWriteOffice\bin\Release')) {
                'Release'
            } else {
                'Debug'
            }
            Add-Type -Path (Join-Path $PSScriptRoot "..\Sources\PSWriteOffice\bin\$configuration\net8.0\DocumentFormat.OpenXml.dll")
        }

        $spreadsheet = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($Path, $true)
        try {
            $slicer = $spreadsheet.WorkbookPart.AddExtendedPart(
                'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
                'application/vnd.ms-excel.slicerCache+xml',
                '.xml')
            $timeline = $spreadsheet.WorkbookPart.AddExtendedPart(
                'http://schemas.microsoft.com/office/2011/relationships/timelineCache',
                'application/vnd.ms-excel.timelineCache+xml',
                '.xml')

            foreach ($part in @($slicer, $timeline)) {
                $stream = $part.GetStream([System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
                try {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes('<root />')
                    $stream.Write($bytes, 0, $bytes.Length)
                } finally {
                    $stream.Dispose()
                }
            }
        } finally {
            $spreadsheet.Dispose()
        }
    }
}

Describe 'Excel DSL surface' {
    It 'creates a workbook with canonical cmdlets' {
        $path = Join-Path $TestDrive 'DslExcel.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
        }

        Test-Path $path | Should -BeTrue

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -BeGreaterThan 0
        } finally {
            Close-OfficeExcel -Document $doc
        }

    }

    It 'creates 1904 date-system workbooks from the thin DSL surface' {
        $path = Join-Path $TestDrive 'DslExcelDateSystem1904.xlsx'
        $date = [datetime] '2024-01-01'

        New-OfficeExcel -Path $path -DateSystem 1904 {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value $date
            }
        }

        Get-TestWorkbookDate1904 -Path $path | Should -Be '1'
        [math]::Abs(([double] (Get-TestWorksheetCellValue -Path $path -Address 'A1')) - (Get-TestExcel1904Serial -Value $date)) | Should -BeLessThan 0.000001

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $summary.DateSystem | Should -Be '1904'
        $summary.Schema.DateSystem | Should -Be '1904'
    }

    It 'exports 1904 date-system workbooks and can set the date system on existing documents' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDateSystem1904.xlsx'
        $date = [datetime] '2024-02-03'

        [PSCustomObject]@{ When = $date } | Export-OfficeExcel -Path $path -DateSystem 1904

        Get-TestWorkbookDate1904 -Path $path | Should -Be '1'
        [math]::Abs(([double] (Get-TestWorksheetCellValue -Path $path -Address 'A2')) - (Get-TestExcel1904Serial -Value $date)) | Should -BeLessThan 0.000001

        $document = Get-OfficeExcel -Path $path
        try {
            $document | Set-OfficeExcelDateSystem -DateSystem 1900 | Out-Null
            $document | Save-OfficeExcel
        } finally {
            $document | Close-OfficeExcel
        }

        $summary = Get-OfficeExcelSummary -Path $path
        $summary.DateSystem | Should -Be '1900'
    }

    It 'does not mutate the in-memory date system when Save-OfficeExcel is invoked with WhatIf' {
        $path = Join-Path $TestDrive 'SaveOfficeExcelDateSystemWhatIf.xlsx'
        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data'
        }

        $document = Get-OfficeExcel -Path $path
        try {
            $before = $document.DateSystem
            $document | Save-OfficeExcel -DateSystem 1904 -WhatIf
            $document.DateSystem | Should -Be $before
        } finally {
            $document | Close-OfficeExcel
        }
    }

    It 'sets workbook theme metadata through the thin command surface' {
        $path = Join-Path $TestDrive 'DslExcelTheme.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Theme'
            }
        }

        $theme = Set-OfficeExcelTheme -Path $path -Default -Name 'Contoso Workbook Theme' -PassThru
        $theme.HasTheme | Should -BeTrue
        $theme.Name | Should -Be 'Contoso Workbook Theme'

        $summary = Get-OfficeExcelSummary -Path $path
        $summary.HasTheme | Should -BeTrue
        $summary.ThemeName | Should -Be 'Contoso Workbook Theme'

        $rename = Set-OfficeExcelTheme -Path $path -Name 'Renamed Workbook Theme' -PassThru
        $rename.HasTheme | Should -BeTrue
        $rename.Name | Should -Be 'Renamed Workbook Theme'
    }

    It 'reports preserved slicer and timeline package parts in workbook summaries' {
        $path = Join-Path $TestDrive 'DslExcelInteractionParts.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Value'
            }
        }

        Add-TestWorkbookInteractionParts -Path $path

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $summary.SlicerPartCount | Should -Be 1
        $summary.TimelinePartCount | Should -Be 1
        $summary.Schema.SlicerPartCount | Should -Be 1
        $summary.Schema.TimelinePartCount | Should -Be 1
        $summary.Schema.HasSlicers | Should -BeTrue
        $summary.Schema.HasTimelines | Should -BeTrue
    }

    It 'authors OfficeIMO-owned slicer and timeline binding metadata through thin commands' {
        $path = Join-Path $TestDrive 'DslExcelSlicerTimelineMetadata.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Region'
            }
        }

        $slicer = Add-OfficeExcelSlicer -Path $path -Name RegionSlicer -SourceName Region -PivotTableName SalesPivot -PassThru
        $timeline = Add-OfficeExcelTimeline -Path $path -Name OrderDateTimeline -SourceName OrderDate -PivotTableName SalesPivot -PassThru

        $slicer.Kind | Should -Be 'Slicer'
        $slicer.ContentType | Should -Be 'application/xml'
        $timeline.Kind | Should -Be 'Timeline'
        $timeline.ContentType | Should -Be 'application/xml'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $summary.SlicerPartCount | Should -Be 1
        $summary.TimelinePartCount | Should -Be 1
        $summary.Schema.HasSlicers | Should -BeFalse
        $summary.Schema.HasTimelines | Should -BeFalse
    }

    It 'copies workbook packages while preserving package interaction parts' {
        $source = Join-Path $TestDrive 'PackageCopySource.xlsx'
        $destination = Join-Path $TestDrive 'PackageCopyDestination.xlsm'

        New-OfficeExcel -Path $source {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address A1 -Value 123
            }
        }

        Add-TestWorkbookInteractionParts -Path $source

        $copy = Copy-OfficeExcelWorkbook -Path $source -DestinationPath $destination -PassThru

        $copy.FullName | Should -Be $destination
        Test-Path $destination | Should -BeTrue
        Get-TestWorksheetCellValue -Path $destination -Address 'A1' | Should -Be '123'

        $summary = Get-OfficeExcelSummary -Path $destination -IncludeSchema
        $summary.SlicerPartCount | Should -Be 1
        $summary.TimelinePartCount | Should -Be 1
        $summary.Schema.HasSlicers | Should -BeTrue
        $summary.Schema.HasTimelines | Should -BeTrue
    }

    It 'joins selected worksheets from another workbook with a stable prefix' {
        $source = Join-Path $TestDrive 'WorkbookJoinSource.xlsx'
        $target = Join-Path $TestDrive 'WorkbookJoinTarget.xlsx'

        New-OfficeExcel -Path $source {
            Add-OfficeExcelSheet -Name 'North' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'North value'
            }
            Add-OfficeExcelSheet -Name 'South' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'South value'
            }
        }

        New-OfficeExcel -Path $target {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Summary'
            }
        }

        $join = Join-OfficeExcelWorkbook -Path $target -SourcePath $source -SourceSheet South -SheetNamePrefix 'Imported '
        $join.SheetCount | Should -Be 1
        $join.SourceSheets | Should -Contain 'South'
        $join.TargetSheets | Should -Contain 'Imported South'

        $summary = Get-OfficeExcelSummary -Path $target -IncludeSheets
        $summary.Sheets.Name | Should -Contain 'Imported South'
        $document = Get-OfficeExcel -Path $target -ReadOnly
        try {
            $importedValue = $null
            $document['Imported South'].TryGetCellText(1, 1, [ref] $importedValue) | Should -BeTrue
            $importedValue | Should -Be 'South value'
        } finally {
            $document | Close-OfficeExcel
        }
    }

    It 'adds connection and query-table package metadata through a thin command' {
        $path = Join-Path $TestDrive 'DslExcelConnectionMetadata.xlsx'
        $connectionXml = '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><connection id="1" name="SalesConnection" type="5" refreshedVersion="7"/></connections>'
        $queryTableXml = '<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="SalesQuery" connectionId="1"/>'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Region'
            }
        }

        $connection = Add-OfficeExcelPackageMetadata -Path $path -Kind Connection -Xml $connectionXml -PassThru
        $queryTable = Add-OfficeExcelPackageMetadata -Path $path -Kind QueryTable -WorksheetName Data -Xml $queryTableXml -PassThru

        $connection.Kind | Should -Be 'Connection'
        $connection.ContentType | Should -Be 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'
        $queryTable.Kind | Should -Be 'QueryTable'
        $queryTable.WorksheetName | Should -Be 'Data'
        $queryTable.ContentType | Should -Be 'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $summary.ConnectionPartCount | Should -Be 1
        $summary.QueryTablePartCount | Should -Be 1
        $summary.Schema.ConnectionPartCount | Should -Be 1
        $summary.Schema.QueryTablePartCount | Should -Be 1
        $summary.Schema.HasConnections | Should -BeTrue
        $summary.Schema.HasQueryTables | Should -BeTrue
    }

    It 'applies reusable print layout presets through a thin command' {
        $path = Join-Path $TestDrive 'DslExcelPrintLayout.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Report' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Region'
                Set-OfficeExcelCell -Address B1 -Value 'Revenue'
                Set-OfficeExcelCell -Address A2 -Value 'EU'
                Set-OfficeExcelCell -Address B2 -Value 100
                Set-OfficeExcelPrintLayout -Preset Report -PrintArea A1:D25 -RepeatFirstColumn 1 -RepeatLastColumn 1
            }
        }

        Get-TestWorksheetPageSetupAttribute -Path $path -Name orientation | Should -Be 'landscape'
        Get-TestWorksheetPageSetupAttribute -Path $path -Name fitToWidth | Should -Be '1'
        Get-TestWorksheetPageSetupAttribute -Path $path -Name fitToHeight | Should -Be '0'
        Get-TestWorksheetPageSetupAttribute -Path $path -Name pageOrder | Should -Be 'downThenOver'
        Get-TestWorksheetPageSetupPropertyAttribute -Path $path -Name fitToPage | Should -Be '1'
        Get-TestWorksheetPageMarginAttribute -Path $path -Name left | Should -Be '0.25'

        Get-TestWorkbookDefinedName -Path $path -Name '_xlnm.Print_Area' | Should -Match '\$A\$1:\$D\$25'
        $titles = Get-TestWorkbookDefinedName -Path $path -Name '_xlnm.Print_Titles'
        $titles | Should -Match '\$1:\$1'
        $titles | Should -Match '\$A:\$A'

        $layoutResult = Set-OfficeExcelPrintLayout -Path $path -Sheet Report -Preset Worksheet -PassThru
        $layoutResult.Path | Should -Be $path
        $layoutResult.SheetName | Should -Be 'Report'
        $layoutResult.Preset | Should -Be 'Worksheet'
    }

    It 'adds dashboard chart presets through a thin command' {
        $path = Join-Path $TestDrive 'DslExcelDashboardChart.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'EU'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'US'; Revenue = 120 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Dashboard' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                $chart = Add-OfficeExcelDashboardChart -TableName Sales -Preset CompactComparison -Row 1 -Column 4 -Title 'Revenue' -PassThru
                $chart.Title | Should -Be 'Revenue'
            }
        } | Out-Null

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.ChartCount | Should -Be 1
        ($summary.Sheets | Where-Object Name -eq 'Dashboard').ChartCount | Should -Be 1

        $chartXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/drawings/charts/chart1.xml'
        $chartXml.OuterXml | Should -Match 'barChart'
        $chartXml.OuterXml | Should -Match 'dLbls'
    }

    It 'builds dashboard tables and charts through a thin command' {
        $path = Join-Path $TestDrive 'DslExcelDashboardBuilder.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'EU'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'US'; Revenue = 120 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Dashboard' -Content {
                $result = $rows | New-OfficeExcelDashboard -Title 'Sales Dashboard' -Subtitle 'Monthly revenue' -TableName 'Sales' -ChartPreset CompactComparison -ChartTitle 'Revenue' -PassThru
                $result.TableRange | Should -Be 'A3:B5'
                $result.TableName | Should -Be 'Sales'
                $result.ChartTitle | Should -Be 'Revenue'
                $result.ChartType | Should -Be 'BarClustered'
            }
        } | Out-Null

        Read-XlsxEntryText -Path $path -Entry 'xl/sharedStrings.xml' | Should -Match 'Sales Dashboard'
        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.ChartCount | Should -Be 1
        ($summary.Sheets | Where-Object Name -eq 'Dashboard').TableCount | Should -Be 1
    }

    It 'keeps dashboard row Document columns as data' {
        $path = Join-Path $TestDrive 'DslExcelDashboardDocumentColumn.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'EU'; Document = 'Proposal'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'US'; Document = 'Contract'; Revenue = 120 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Dashboard' -Content {
                $result = $rows | New-OfficeExcelDashboard -Title 'Document Dashboard' -TableName 'Documents' -NoChart -PassThru
                $result.TableRange | Should -Be 'A3:C5'
                $result.TableName | Should -Be 'Documents'
            }
        } | Out-Null

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName Dashboard -Range 'A3:C5')
        $rows.Count | Should -Be 2
        $rows[0].Document | Should -Be 'Proposal'
        $rows[1].Document | Should -Be 'Contract'
    }

    It 'runs workbook preflight checks through the reusable OfficeIMO report' {
        $path = Join-Path $TestDrive 'DslExcelPreflight.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Region'
                Set-OfficeExcelCell -Address B1 -Value 'Revenue'
                Set-OfficeExcelCell -Address A2 -Value 'EU'
                Set-OfficeExcelCell -Address B2 -Value 100
            }
        } | Out-Null

        $preflight = Get-OfficeExcelPreflight -Path $path -Capability ReadWorkbookData,EditCellValues -IncludeFeatures
        $preflight.HasAdvancedFeatures | Should -BeFalse
        $preflight.FeatureCount | Should -BeGreaterThan 0
        @($preflight.Capabilities | Where-Object Name -eq 'ReadWorkbookData')[0].CanAttempt | Should -BeTrue
        @($preflight.Capabilities | Where-Object Name -eq 'EditCellValues')[0].CanAttempt | Should -BeTrue
        @($preflight.Features | Where-Object Name -eq 'Worksheets')[0].SupportLevel | Should -Be 'Editable'

        $markdown = Get-OfficeExcelPreflight -Path $path -AsMarkdown
        $markdown | Should -Match 'Capability Preflight'

        { Get-OfficeExcelPreflight -Path $path -Capability ReadWorkbookData -ThrowOnFailure -ErrorAction Stop } | Should -Not -Throw
    }

    It 'returns preflight repair hints through the thin command surface' {
        $path = Join-Path $TestDrive 'DslExcelPreflightRepairHints.xlsx'

        New-OfficeExcel -Path $path -ForceFullCalculationOnOpen {
            Add-OfficeExcelSheet -Name 'Calc' -Content {
                Set-OfficeExcelCell -Address A1 -Value 2
                Set-OfficeExcelFormula -Address B1 -Formula 'A1+1'
            }
        } | Out-Null

        $preflight = Get-OfficeExcelPreflight -Path $path -Capability UseCachedFormulaValues -IncludeRepairHints
        $capability = @($preflight.Capabilities | Where-Object Name -eq 'UseCachedFormulaValues')[0]
        $capability.CanAttempt | Should -BeFalse
        @($capability.RepairHints | Where-Object FeatureName -eq 'Missing formula caches').Count | Should -BeGreaterThan 0
        ($capability.RepairHints | Select-Object -First 1).Action | Should -Match 'Refresh cached formula values'

        $markdown = Get-OfficeExcelPreflight -Path $path -AsMarkdown
        $markdown | Should -Match 'Repair Hints'
    }

    It 'creates workbooks from package-preserving templates' {
        $template = Join-Path $TestDrive 'TemplateSource.xlsx'
        $path = Join-Path $TestDrive 'TemplateOutput.xlsx'

        New-OfficeExcel -Path $template {
            Add-OfficeExcelSheet -Name 'Template' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 42 -BackgroundColor '#D9EAD3'
            }
        }

        New-OfficeExcel -TemplatePath $template -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 100
            }
        }

        Get-TestWorksheetCellValue -Path $path -Address 'A1' | Should -Be '42'
        Get-TestWorksheetCellStyleIndex -Path $path -Address 'A1' | Should -Be (Get-TestWorksheetCellStyleIndex -Path $template -Address 'A1')
        Get-TestWorksheetCellValue -Path $path -Address 'A1' -Entry 'xl/worksheets/sheet2.xml' | Should -Be '100'
    }

    It 'supports transposed Excel tables' {
        $path = Join-Path $TestDrive 'TransposedExcelTable.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'Europe'; Revenue = 21704714 }
            [PSCustomObject]@{ Region = 'Asia'; Revenue = 8774099 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -View Transpose -TableName 'TransposedSales'
            }
        }

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:C3')
        $imported[0].Property | Should -Be 'Region'
        $imported[0].Row1 | Should -Be 'Europe'
        $imported[0].Row2 | Should -Be 'Asia'
        $imported[1].Property | Should -Be 'Revenue'
        $imported[1].Row1 | Should -Be 21704714
        $imported[1].Row2 | Should -Be 8774099
    }

    It 'supports transposed Excel tables from IDataReader input' {
        $path = Join-Path $TestDrive 'TransposedExcelReaderTable.xlsx'
        $table = [System.Data.DataTable]::new('SqlRows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('One', 1)
        [void] $table.Rows.Add('Two', 2)
        $reader = $table.CreateDataReader()

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $reader -View Transpose -TableName 'TransposedReader'
            }
        }

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:C3')
        $imported[0].Property | Should -Be 'Name'
        $imported[0].Row1 | Should -Be 'One'
        $imported[0].Row2 | Should -Be 'Two'
        $imported[1].Property | Should -Be 'Value'
        $imported[1].Row1 | Should -Be 1
        $imported[1].Row2 | Should -Be 2
    }

    It 'applies table visual style flags from friendly switches' {
        $path = Join-Path $TestDrive 'TableVisualFlags.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -ShowFirstColumn -ShowLastColumn -NoRowStripes -ShowColumnStripes
            }
        }

        $tableXml = Read-XlsxEntryText -Path $path -Entry 'xl/tables/table1.xml'
        $tableXml | Should -Match 'showFirstColumn="(?:1|true)"'
        $tableXml | Should -Match 'showLastColumn="(?:1|true)"'
        $tableXml | Should -Match 'showRowStripes="(?:0|false)"'
        $tableXml | Should -Match 'showColumnStripes="(?:1|true)"'

        $falseSwitchPath = Join-Path $TestDrive 'TableVisualFalseSwitches.xlsx'
        $falseSwitchParameters = @{
            InputObject = $rows
            TableName = 'Sales'
            ShowFirstColumn = $false
            NoRowStripes = $false
            ShowColumnStripes = $false
        }
        New-OfficeExcel -Path $falseSwitchPath {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable @falseSwitchParameters
            }
        }

        $falseSwitchTableXml = Read-XlsxEntryText -Path $falseSwitchPath -Entry 'xl/tables/table1.xml'
        $falseSwitchTableXml | Should -Not -Match 'showFirstColumn="(?:1|true)"'
        $falseSwitchTableXml | Should -Not -Match 'showColumnStripes="(?:1|true)"'
        $falseSwitchTableXml | Should -Not -Match 'showRowStripes="(?:0|false)"'

        $exportPath = Join-Path $TestDrive 'ExportTableVisualFlags.xlsx'
        $rows | Export-OfficeExcel -Path $exportPath -TableName 'ExportedSales' -ShowFirstColumn -ShowLastColumn -NoRowStripes -ShowColumnStripes
        $exportTableXml = Read-XlsxEntryText -Path $exportPath -Entry 'xl/tables/table1.xml'
        $exportTableXml | Should -Match 'showFirstColumn="(?:1|true)"'
        $exportTableXml | Should -Match 'showLastColumn="(?:1|true)"'
        $exportTableXml | Should -Match 'showRowStripes="(?:0|false)"'
        $exportTableXml | Should -Match 'showColumnStripes="(?:1|true)"'

        $dataSetPath = Join-Path $TestDrive 'DataSetTableVisualFlags.xlsx'
        $dataSet = [System.Data.DataSet]::new('Book')
        $dataTable = [System.Data.DataTable]::new('Data')
        [void] $dataTable.Columns.Add('Region', [string])
        [void] $dataTable.Columns.Add('Sales', [int])
        [void] $dataTable.Rows.Add('NA', 100)
        [void] $dataTable.Rows.Add('EMEA', 200)
        [void] $dataSet.Tables.Add($dataTable)
        New-OfficeExcel -Path $dataSetPath {
            Add-OfficeExcelDataSet -DataSet $dataSet -ShowFirstColumn -ShowLastColumn -NoRowStripes -ShowColumnStripes
        }

        $dataSetTableXml = Read-XlsxEntryText -Path $dataSetPath -Entry 'xl/tables/table1.xml'
        $dataSetTableXml | Should -Match 'showFirstColumn="(?:1|true)"'
        $dataSetTableXml | Should -Match 'showLastColumn="(?:1|true)"'
        $dataSetTableXml | Should -Match 'showRowStripes="(?:0|false)"'
        $dataSetTableXml | Should -Match 'showColumnStripes="(?:1|true)"'

        $reportPath = Join-Path $TestDrive 'ReportTableVisualFlags.xlsx'
        New-OfficeExcel -Path $reportPath {
            Add-OfficeExcelReportSheet -Name 'Summary' {
                Add-OfficeExcelReportTable -InputObject $rows -Title 'Sales' -ShowFirstColumn -ShowLastColumn -NoRowStripes -ShowColumnStripes
            }
        }

        $reportTableXml = Read-XlsxEntryText -Path $reportPath -Entry 'xl/tables/table1.xml'
        $reportTableXml | Should -Match 'showFirstColumn="(?:1|true)"'
        $reportTableXml | Should -Match 'showLastColumn="(?:1|true)"'
        $reportTableXml | Should -Match 'showRowStripes="(?:0|false)"'
        $reportTableXml | Should -Match 'showColumnStripes="(?:1|true)"'
    }

    It 'round-trips encrypted workbooks through lifecycle cmdlets' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.Excel.ExcelDocument' -MethodName 'LoadEncrypted')) {
            (Get-Command New-OfficeExcel).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Save-OfficeExcel).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Get-OfficeExcel).Parameters.Keys | Should -Contain 'Password'
            return
        }

        $path = Join-Path $TestDrive 'EncryptedExcel.xlsx'

        New-OfficeExcel -Path $path -Password 'secret' -SafePreflight {
            Set-OfficeExcelExecutionPolicy -Mode Sequential -ParallelThreshold 5 -WorksheetValidation Always -Diagnostics
            Add-OfficeExcelSheet -Name 'Secure' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Encrypted value'
            }
        }

        { Get-ZipEntriesLocal -Path $path } | Should -Throw
        $doc = Get-OfficeExcel -Path $path -Password 'secret' -ReadOnly
        try {
            $doc.Sheets[0].Name | Should -Be 'Secure'
            $value = $null
            $doc.Sheets[0].TryGetCellText(1, 1, [ref] $value) | Should -BeTrue
            $value | Should -Be 'Encrypted value'
            $summary = Get-OfficeExcelSummary -Document $doc -IncludeSheets -IncludeSchema
            $summary.SheetCount | Should -Be 1
            $summary.Sheets[0].Name | Should -Be 'Secure'
            $summary.Schema.Worksheets[0].Name | Should -Be 'Secure'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $doc = Get-OfficeExcel -Path $path -Password 'secret'
        try {
            { $doc | Save-OfficeExcel -Path $path -ErrorAction Stop } |
                Should -Throw '*Provide -Password*'
            $doc.Sheets[0].Cell(1, 1, 'Must not be saved without a password', $null, $null)
            { $doc | Close-OfficeExcel -Save -ErrorAction Stop } |
                Should -Throw '*Provide -Password*'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $plainCopyPath = Join-Path $TestDrive 'EncryptedExcelPlainCopy.xlsx'
        $doc = Get-OfficeExcel -Path $path -Password 'secret'
        try {
            $doc | Save-OfficeExcel -Path $plainCopyPath
            $doc.Sheets[0].Cell(1, 1, 'Updated plain copy', $null, $null)
            $doc | Save-OfficeExcel
        } finally {
            Close-OfficeExcel -Document $doc
        }

        { Get-ZipEntriesLocal -Path $plainCopyPath } | Should -Not -Throw
        $plainCopy = Get-OfficeExcel -Path $plainCopyPath -ReadOnly
        try {
            $plainValue = $null
            $plainCopy.Sheets[0].TryGetCellText(1, 1, [ref] $plainValue) | Should -BeTrue
            $plainValue | Should -Be 'Updated plain copy'
        } finally {
            Close-OfficeExcel -Document $plainCopy
        }

        $encryptedCopyPath = Join-Path $TestDrive 'EncryptedExcelCopy.xlsx'
        $doc = Get-OfficeExcel -Path $path -Password 'secret'
        try {
            $doc | Save-OfficeExcel -Path $encryptedCopyPath -Password 'copy-secret'
            $doc.Sheets[0].Cell(1, 1, 'Updated encrypted copy', $null, $null)
            { $doc | Save-OfficeExcel -ErrorAction Stop } | Should -Throw '*Provide -Password*'
            $doc | Save-OfficeExcel -Password 'copy-secret'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $encryptedCopy = Get-OfficeExcel -Path $encryptedCopyPath -Password 'copy-secret' -ReadOnly
        try {
            $copyValue = $null
            $encryptedCopy.Sheets[0].TryGetCellText(1, 1, [ref] $copyValue) | Should -BeTrue
            $copyValue | Should -Be 'Updated encrypted copy'
        } finally {
            Close-OfficeExcel -Document $encryptedCopy
        }

        $doc = Get-OfficeExcel -Path $path -Password 'secret'
        try {
            $doc.Sheets[0].Cell(1, 1, 'Updated encrypted value', $null, $null)
            $doc | Save-OfficeExcel -Password 'secret'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $doc = Get-OfficeExcel -Path $path -Password 'secret' -ReadOnly
        try {
            $value = $null
            $doc.Sheets[0].TryGetCellText(1, 1, [ref] $value) | Should -BeTrue
            $value | Should -Be 'Updated encrypted value'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        (Get-Command Get-OfficeExcel).Parameters.Keys | Should -Not -Contain 'AutoSave'
    }

    It 'clears encrypted workbook associations after HTML conversion' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.Excel.ExcelDocument' -MethodName 'LoadEncrypted')) {
            (Get-Command ConvertTo-OfficeExcelHtml).Parameters.Keys | Should -Contain 'Password'
            return
        }

        $path = Join-Path $TestDrive 'EncryptedExcelHtml.xlsx'
        $htmlPath = Join-Path $TestDrive 'EncryptedExcelHtml.html'
        New-OfficeExcel -Path $path -Password 'secret' {
            Add-OfficeExcelSheet -Name 'Secure' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Encrypted HTML value'
            }
        }

        $assembly = (Get-Command Get-OfficeExcel).ImplementingType.Assembly
        $flags = [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static
        $serviceType = $assembly.GetType('PSWriteOffice.Services.Excel.ExcelDocumentService')
        $associatedPaths = $serviceType.GetField('AssociatedPaths', $flags)
        $encryptedPaths = $serviceType.GetField('EncryptedSourcePaths', $flags)
        $associatedCount = $associatedPaths.GetValue($null).Count
        $encryptedCount = $encryptedPaths.GetValue($null).Count

        ConvertTo-OfficeExcelHtml -Path $path -Password 'secret' -OutputPath $htmlPath

        Test-Path -LiteralPath $htmlPath | Should -BeTrue
        $associatedPaths.GetValue($null).Count | Should -Be $associatedCount
        $encryptedPaths.GetValue($null).Count | Should -Be $encryptedCount
    }

    It 'configures Excel execution policy from PowerShell' {
        $path = Join-Path $TestDrive 'ExcelExecutionPolicy.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Policy'
        }
        $doc = Get-OfficeExcel -Path $path
        try {
            $result = $doc | Set-OfficeExcelExecutionPolicy -Mode Parallel -ParallelThreshold 3 -MaxDegreeOfParallelism 2 -WorksheetValidation Disabled -Diagnostics -DisableAutoFitImmediateSave -PassThru

            $result | Should -Be $doc
            $doc.Execution.Mode.ToString() | Should -Be 'Parallel'
            $doc.Execution.ParallelThreshold | Should -Be 3
            $doc.Execution.MaxDegreeOfParallelism | Should -Be 2
            $doc.Execution.WorksheetValidation.ToString() | Should -Be 'Disabled'
            $doc.Execution.DiagnosticsRequested | Should -BeTrue
            $doc.Execution.SaveWorksheetAfterAutoFit | Should -BeFalse
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'supports alias-only syntax' {
        $path = Join-Path $TestDrive 'DslExcelAlias.xlsx'
        $rows = @(
            [PSCustomObject]@{ Item = 'Laptop'; Qty = 5 }
            [PSCustomObject]@{ Item = 'Tablet'; Qty = 12 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Orders' -Content {
                ExcelCell -Address 'A1' -Value 'Item'
                ExcelCell -Address 'B1' -Value 'Qty'
                ExcelTable -InputObject $rows -TableName 'OrdersTable'
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'exports workbook merge aliases from the manifest' {
        (Get-Alias -Name Merge-OfficeExcelWorkbook -ErrorAction Stop).ResolvedCommandName | Should -Be 'Join-OfficeExcelWorkbook'
        (Get-Alias -Name ExcelWorkbookJoin -ErrorAction Stop).ResolvedCommandName | Should -Be 'Join-OfficeExcelWorkbook'
        (Get-Alias -Name ExcelWorkbookMerge -ErrorAction Stop).ResolvedCommandName | Should -Be 'Join-OfficeExcelWorkbook'
    }

    It 'preserves legacy Excel table data parameter aliases' {
        $path = Join-Path $TestDrive 'DslExcelTableDataAliases.xlsx'
        $rows = @(
            [PSCustomObject]@{ Item = 'Laptop'; Qty = 5 }
            [PSCustomObject]@{ Item = 'Tablet'; Qty = 12 }
        )

        $table = [System.Data.DataTable]::new('Stock')
        [void] $table.Columns.Add('Item', [string])
        [void] $table.Columns.Add('Qty', [int])
        [void] $table.Rows.Add('Dock', 3)

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'DataAlias' -Content {
                ExcelTable -Data $rows -TableName 'RowsAlias'
            }
            Add-OfficeExcelSheet -Name 'DataTableAlias' -Content {
                ExcelTable -DataTable $table -TableName 'TableAlias'
            }
        }

        $rowsAlias = @(Import-OfficeExcel -Path $path -WorksheetName 'DataAlias' -Range 'A1:B3')
        $tableAlias = @(Import-OfficeExcel -Path $path -WorksheetName 'DataTableAlias' -Range 'A1:B2')
        $rowsAlias[0].Item | Should -Be 'Laptop'
        $rowsAlias[1].Qty | Should -Be 12
        $tableAlias[0].Item | Should -Be 'Dock'
        $tableAlias[0].Qty | Should -Be 3
    }

    It 'writes a DataTable directly as an Excel table' {
        $path = Join-Path $TestDrive 'DslExcelDataTable.xlsx'
        $table = [System.Data.DataTable]::new('People')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Score', [int])
        [void] $table.Rows.Add('Ada', 10)
        [void] $table.Rows.Add('Grace', 20)

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $table -TableName 'PeopleTable' -AutoFit
            }
        }

        Test-Path $path | Should -BeTrue
        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'PeopleTable')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A1:B3'

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[0].Name | Should -Be 'Ada'
        $imported[0].Score | Should -Be 10
    }

    It 'appends rows to an existing Excel table outside the DSL' {
        $path = Join-Path $TestDrive 'ExcelExistingTableAppend.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
        }

        $doc = Get-OfficeExcel -Path $path
        try {
            $table = $doc | Add-OfficeExcelTableRow -Sheet Data -TableName Sales -InputObject ([pscustomobject]@{ Region = 'APAC'; Revenue = 300 }) -PassThru
            $table.Range | Should -Be 'A1:B4'

            Close-OfficeExcel -Document $doc -Save
            $doc = $null
        } finally {
            if ($null -ne $doc) {
                Close-OfficeExcel -Document $doc
            }
        }

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B4')
        $imported.Count | Should -Be 3
        $imported[2].Region | Should -Be 'APAC'
        $imported[2].Revenue | Should -Be 300
    }

    It 'finds a named Excel table on later sheets when appending without a sheet filter' {
        $path = Join-Path $TestDrive 'ExcelExistingTableAppendWithoutSheet.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Overview'
            }
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
        }

        Add-OfficeExcelTableRow -Path $path -TableName Sales -InputObject ([pscustomobject]@{ Region = 'APAC'; Revenue = 300 })

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[1].Region | Should -Be 'APAC'
        $imported[1].Revenue | Should -Be 300
    }

    It 'does not emit a live table from path-owned Excel table appends' {
        $path = Join-Path $TestDrive 'ExcelPathAppendPassThru.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
        }

        $warnings = @()
        $result = Add-OfficeExcelTableRow -Path $path -TableName Sales -InputObject ([pscustomobject]@{ Region = 'APAC'; Revenue = 300 }) -PassThru -WarningVariable warnings

        $result | Should -BeNullOrEmpty
        $warnings[0].Message | Should -BeLike '*no live ExcelTable is emitted*'
        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[1].Region | Should -Be 'APAC'
        $imported[1].Revenue | Should -Be 300
    }

    It 'appends explicit input to each piped Excel table target' {
        $path1 = Join-Path $TestDrive 'ExcelPipedTableAppend1.xlsx'
        $path2 = Join-Path $TestDrive 'ExcelPipedTableAppend2.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
        )

        foreach ($path in @($path1, $path2)) {
            New-OfficeExcel -Path $path {
                Add-OfficeExcelSheet -Name 'Data' -Content {
                    Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                }
            }
        }

        $doc1 = Get-OfficeExcel -Path $path1
        $doc2 = Get-OfficeExcel -Path $path2
        try {
            $table1 = $doc1.Sheets[0].Table('Sales')
            $table2 = $doc2.Sheets[0].Table('Sales')

            @($table1, $table2) |
                Add-OfficeExcelTableRow -InputObject ([pscustomobject]@{ Region = 'APAC'; Revenue = 300 })

            Close-OfficeExcel -Document $doc1 -Save
            Close-OfficeExcel -Document $doc2 -Save
            $doc1 = $null
            $doc2 = $null
        } finally {
            if ($null -ne $doc1) {
                Close-OfficeExcel -Document $doc1
            }
            if ($null -ne $doc2) {
                Close-OfficeExcel -Document $doc2
            }
        }

        foreach ($path in @($path1, $path2)) {
            $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
            $imported.Count | Should -Be 2
            $imported[1].Region | Should -Be 'APAC'
            $imported[1].Revenue | Should -Be 300
        }
    }

    It 'writes a DataSet as one worksheet per table' {
        $path = Join-Path $TestDrive 'DslExcelDataSet.xlsx'
        $dataSet = [System.Data.DataSet]::new('Report')

        $sales = [System.Data.DataTable]::new('Sales:2026')
        [void] $sales.Columns.Add('Region', [string])
        [void] $sales.Columns.Add('Revenue', [int])
        [void] $sales.Rows.Add('NA', 100)
        [void] $sales.Rows.Add('EMEA', 200)
        [void] $dataSet.Tables.Add($sales)

        $notes = [System.Data.DataTable]::new('Notes')
        [void] $notes.Columns.Add('Text', [string])
        [void] $notes.Rows.Add('Checked')
        [void] $dataSet.Tables.Add($notes)

        New-OfficeExcel -Path $path {
            Add-OfficeExcelDataSet -DataSet $dataSet -AutoFit
        }

        Test-Path $path | Should -BeTrue
        $salesRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B3')
        $salesRows.Count | Should -Be 2
        $salesRows[1].Region | Should -Be 'EMEA'
        $salesRows[1].Revenue | Should -Be 200

        $notesRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Notes' -Range 'A1:A2')
        $notesRows.Count | Should -Be 1
        $notesRows[0].Text | Should -Be 'Checked'
    }

    It 'exports and imports objects through operator cmdlets' {
        $path = Join-Path $TestDrive 'ExportOfficeExcel.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100; Internal = 'skip' }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200; Internal = 'skip' }
        )

        $file = $rows |
            Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -Title 'Sales Export' -AutoFit -FreezeTopRow -BoldTopRow -ExcludeProperty Internal -PassThru

        $file.FullName | Should -Be $path
        Test-Path $path | Should -BeTrue

        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A2:B4'

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A2:B4')
        $imported.Count | Should -Be 2
        $imported[0].Region | Should -Be 'NA'
        $imported[0].Revenue | Should -Be 100
        $imported[0].PSObject.Properties.Name | Should -Not -Contain 'Internal'

        { Import-OfficeExcel -Path $path -WorksheetName 'Data' -StartRow 4 -EndRow 2 -StartColumn 1 -EndColumn 2 } |
            Should -Throw '*StartRow must be less than or equal to EndRow*'
        { Import-OfficeExcel -Path $path -WorksheetName 'Data' -StartRow 2 -EndRow 4 -StartColumn 3 -EndColumn 1 } |
            Should -Throw '*StartColumn must be less than or equal to EndColumn*'
    }

    It 'exports plain objects through the default table path' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDefaultObjects.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100; Created = [DateTime] '2026-01-01'; Enabled = $true }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200; Created = [DateTime] '2026-01-02'; Enabled = $false }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales'

        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A1:D3'

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:D3')
        $imported.Count | Should -Be 2
        $imported[0].Region | Should -Be 'NA'
        $imported[0].Revenue | Should -Be 100
        $imported[1].Region | Should -Be 'EMEA'
        $imported[1].Revenue | Should -Be 200
        $imported[1].Enabled | Should -BeFalse
    }

    It 'allows NoClobber as a safety option when appending Excel rows' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelAppendNoClobber.xlsx'

        [PSCustomObject]@{ Region = 'NA'; Revenue = 100 } |
            Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales'

        [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 } |
            Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -Append -NoClobber

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data')
        $imported.Count | Should -Be 2
        $imported[0].Region | Should -Be 'NA'
        $imported[1].Region | Should -Be 'EMEA'
        $imported[1].Revenue | Should -Be 200
    }

    It 'allows NoClobber as a safety option when clearing an Excel sheet' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelClearSheetNoClobber.xlsx'

        [PSCustomObject]@{ Region = 'NA'; Revenue = 100 } |
            Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales'

        [PSCustomObject]@{ Region = 'APAC'; Revenue = 300 } |
            Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -ClearSheet -NoClobber

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data')
        $imported.Count | Should -Be 1
        $imported[0].Region | Should -Be 'APAC'
        $imported[0].Revenue | Should -Be 300
    }

    It 'applies export-time column formats by header' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelColumnFormats.xlsx'
        $rows = @(
            [PSCustomObject]@{ Id = '00042'; Revenue = 1234.5; Rate = 0.125; Created = [DateTime] '2026-06-23' }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -TextColumn Id -CurrencyColumn Revenue -ColumnFormat @{
            Rate = @{ Style = 'Percent'; Decimals = 1 }
            Created = 'Date'
        } -FormatCultureName en-US -AutoFitFormattedColumn

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $idCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='A2']")
        $revenueCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='B2']")
        $rateCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='C2']")
        $createdCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='D2']")

        $idCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $revenueCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $rateCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $createdCell.GetAttribute('s') | Should -Not -BeNullOrEmpty

        $stylesXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/styles.xml'
        $formats = @($stylesXml.SelectNodes("//*[local-name()='numFmt']") | ForEach-Object { $_.GetAttribute('formatCode') })
        $formats | Should -Contain '@'
        ($formats -join '|') | Should -Match '\$'
        ($formats -join '|') | Should -Match '0\.0%'
        ($formats -join '|') | Should -Match 'yyyy-mm-dd'

        { $rows | Export-OfficeExcel -Path (Join-Path $TestDrive 'ExportOfficeExcelMissingColumnFormat.xlsx') -ColumnFormat @{ Missing = 'Integer' } -ErrorAction Stop } |
            Should -Throw '*Column format headers were not found*'

        { $rows | Export-OfficeExcel -Path (Join-Path $TestDrive 'ExportOfficeExcelNoHeaderColumnFormat.xlsx') -NoHeader -CurrencyColumn Revenue -ErrorAction Stop } |
            Should -Throw '*require a header row*'

        $rawAppendPath = Join-Path $TestDrive 'ExportOfficeExcelRawAppendMissingHeader.xlsx'
        @([PSCustomObject]@{ Region = 'Revenue'; Amount = 100 }) |
            Export-OfficeExcel -Path $rawAppendPath -WorksheetName 'Data' -NoTable
        { [PSCustomObject]@{ Region = 'EMEA'; Amount = 200 } | Export-OfficeExcel -Path $rawAppendPath -WorksheetName 'Data' -Append -NoTable -CurrencyColumn Revenue -ErrorAction Stop } |
            Should -Throw '*Column format headers were not found*'

        $rawAppendFormatPath = Join-Path $TestDrive 'ExportOfficeExcelRawAppendColumnFormats.xlsx'
        @([PSCustomObject]@{ Region = 'NA'; Revenue = 123.45 }) |
            Export-OfficeExcel -Path $rawAppendFormatPath -WorksheetName 'Data' -NoTable
        [PSCustomObject]@{ Region = 'EMEA'; Revenue = 987.65 } |
            Export-OfficeExcel -Path $rawAppendFormatPath -WorksheetName 'Data' -Append -NoTable -CurrencyColumn Revenue -FormatCultureName en-US
        $rawAppendSheetXml = Get-ZipXmlDocumentLocal -Path $rawAppendFormatPath -Entry 'xl/worksheets/sheet1.xml'
        $rawAppendRevenueCell = $rawAppendSheetXml.SelectSingleNode("//*[local-name()='c' and @r='B3']")
        $rawAppendRevenueCell.GetAttribute('s') | Should -Not -BeNullOrEmpty

        $rawAppendPartialFormatPath = Join-Path $TestDrive 'ExportOfficeExcelRawAppendPartialColumnFormats.xlsx'
        @([PSCustomObject]@{ Region = 'NA'; Revenue = 123.45 }) |
            Export-OfficeExcel -Path $rawAppendPartialFormatPath -WorksheetName 'Data' -NoTable
        [PSCustomObject]@{ Region = 'EMEA'; Revenue = 987.65 } |
            Export-OfficeExcel -Path $rawAppendPartialFormatPath -WorksheetName 'Data' -Append -NoTable -CurrencyColumn Revenue -IntegerColumn Missing -IgnoreMissingColumnFormat -FormatCultureName en-US
        $rawAppendPartialSheetXml = Get-ZipXmlDocumentLocal -Path $rawAppendPartialFormatPath -Entry 'xl/worksheets/sheet1.xml'
        $rawAppendPartialRevenueCell = $rawAppendPartialSheetXml.SelectSingleNode("//*[local-name()='c' and @r='B3']")
        $rawAppendPartialRevenueCell.GetAttribute('s') | Should -Not -BeNullOrEmpty

        $headerlessTableAppendPath = Join-Path $TestDrive 'ExportOfficeExcelHeaderlessTableAppendColumnFormats.xlsx'
        @([PSCustomObject]@{ Region = 'Revenue'; Amount = 100 }) |
            Export-OfficeExcel -Path $headerlessTableAppendPath -WorksheetName 'Data' -TableName 'Sales' -NoHeader
        { [PSCustomObject]@{ Region = 'EMEA'; Amount = 200 } | Export-OfficeExcel -Path $headerlessTableAppendPath -WorksheetName 'Data' -TableName 'Sales' -Append -NoHeader -CurrencyColumn Revenue -ErrorAction Stop } |
            Should -Throw '*require a header row*'

        $headerlessRawAppendPath = Join-Path $TestDrive 'ExportOfficeExcelHeaderlessRawAppendColumnFormats.xlsx'
        @([PSCustomObject]@{ Region = 'Revenue'; Amount = 100 }) |
            Export-OfficeExcel -Path $headerlessRawAppendPath -WorksheetName 'Data' -NoTable -NoHeader
        { [PSCustomObject]@{ Region = 'EMEA'; Amount = 200 } | Export-OfficeExcel -Path $headerlessRawAppendPath -WorksheetName 'Data' -Append -NoTable -NoHeader -CurrencyColumn Revenue -ErrorAction Stop } |
            Should -Throw '*require a header row*'

        $titlePath = Join-Path $TestDrive 'ExportOfficeExcelColumnFormatsWithTitle.xlsx'
        $rows | Export-OfficeExcel -Path $titlePath -WorksheetName 'Data' -TableName 'Sales' -Title 'Sales Export' -CurrencyColumn Revenue -FormatCultureName en-US
        $titleSheetXml = Get-ZipXmlDocumentLocal -Path $titlePath -Entry 'xl/worksheets/sheet1.xml'
        $titleRevenueCell = $titleSheetXml.SelectSingleNode("//*[local-name()='c' and @r='B3']")
        $titleRevenueCell.GetAttribute('s') | Should -Not -BeNullOrEmpty

        $appendTitlePath = Join-Path $TestDrive 'ExportOfficeExcelAppendColumnFormatsWithTitle.xlsx'
        $rows | Export-OfficeExcel -Path $appendTitlePath -WorksheetName 'Data' -TableName 'Sales' -Title 'Sales Export'
        [PSCustomObject]@{ Id = '00043'; Revenue = 987.65; Rate = 0.2; Created = [DateTime] '2026-06-24' } |
            Export-OfficeExcel -Path $appendTitlePath -WorksheetName 'Data' -TableName 'Sales' -Append -CurrencyColumn Revenue -FormatCultureName en-US
        $appendTitleSheetXml = Get-ZipXmlDocumentLocal -Path $appendTitlePath -Entry 'xl/worksheets/sheet1.xml'
        $appendTitleRevenueCell = $appendTitleSheetXml.SelectSingleNode("//*[local-name()='c' and @r='B4']")
        $appendTitleRevenueCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $appendTitleStylesXml = Get-ZipXmlDocumentLocal -Path $appendTitlePath -Entry 'xl/styles.xml'
        $appendTitleFormats = @($appendTitleStylesXml.SelectNodes("//*[local-name()='numFmt']") | ForEach-Object { $_.GetAttribute('formatCode') })
        ($appendTitleFormats -join '|') | Should -Match '\$'

        $customPath = Join-Path $TestDrive 'ExportOfficeExcelNumericCustomFormat.xlsx'
        [PSCustomObject]@{ Count = 42 } | Export-OfficeExcel -Path $customPath -ColumnFormat @{ Count = '00000' }
        $customStylesXml = Get-ZipXmlDocumentLocal -Path $customPath -Entry 'xl/styles.xml'
        $customFormats = @($customStylesXml.SelectNodes("//*[local-name()='numFmt']") | ForEach-Object { $_.GetAttribute('formatCode') })
        $customFormats | Should -Contain '00000'

        $friendlyPresetPath = Join-Path $TestDrive 'ExportOfficeExcelFriendlyPresetColumnFormats.xlsx'
        [PSCustomObject]@{ Created = [DateTime] '2026-06-24'; Updated = [DateTime] '2026-06-25' } |
            Export-OfficeExcel -Path $friendlyPresetPath -ColumnFormat @{ Created = 'Date Time'; Updated = @{ Style = 'Date-Time' } }
        $friendlyPresetStylesXml = Get-ZipXmlDocumentLocal -Path $friendlyPresetPath -Entry 'xl/styles.xml'
        $friendlyPresetFormats = @($friendlyPresetStylesXml.SelectNodes("//*[local-name()='numFmt']") | ForEach-Object { $_.GetAttribute('formatCode') })
        ($friendlyPresetFormats -join '|') | Should -Match 'yyyy-mm-dd hh:mm:ss'
        $friendlyPresetFormats | Should -Not -Contain 'Date Time'
        $friendlyPresetFormats | Should -Not -Contain 'Date-Time'

        $appendNoHeaderExistingTablePath = Join-Path $TestDrive 'ExportOfficeExcelAppendNoHeaderExistingTableColumnFormats.xlsx'
        @([PSCustomObject]@{ Region = 'NA'; Revenue = 123.45 }) |
            Export-OfficeExcel -Path $appendNoHeaderExistingTablePath -WorksheetName 'Data' -TableName 'Sales'
        [PSCustomObject]@{ Region = 'EMEA'; Revenue = 987.65 } |
            Export-OfficeExcel -Path $appendNoHeaderExistingTablePath -WorksheetName 'Data' -TableName 'Sales' -Append -NoHeader -CurrencyColumn Revenue -FormatCultureName en-US
        $appendNoHeaderExistingTableSheetXml = Get-ZipXmlDocumentLocal -Path $appendNoHeaderExistingTablePath -Entry 'xl/worksheets/sheet1.xml'
        $appendNoHeaderExistingTableRevenueCell = $appendNoHeaderExistingTableSheetXml.SelectSingleNode("//*[local-name()='c' and @r='B3']")
        $appendNoHeaderExistingTableRevenueCell.GetAttribute('s') | Should -Not -BeNullOrEmpty

        $dataSet = [System.Data.DataSet]::new('Mixed')
        $sales = [System.Data.DataTable]::new('Sales')
        $null = $sales.Columns.Add('Region', [string])
        $null = $sales.Columns.Add('Revenue', [decimal])
        $null = $sales.Rows.Add('NA', 100.5)
        $inventory = [System.Data.DataTable]::new('Inventory')
        $null = $inventory.Columns.Add('Item', [string])
        $null = $inventory.Columns.Add('Count', [int])
        $null = $inventory.Rows.Add('Widget', 3)
        $dataSet.Tables.Add($sales)
        $dataSet.Tables.Add($inventory)

        { $dataSet | Export-OfficeExcel -Path (Join-Path $TestDrive 'ExportOfficeExcelDataSetColumnFormats.xlsx') -CurrencyColumn Revenue -IntegerColumn Count -ErrorAction Stop } |
            Should -Not -Throw
        { $dataSet | Export-OfficeExcel -Path (Join-Path $TestDrive 'ExportOfficeExcelDataSetMissingColumnFormat.xlsx') -ColumnFormat @{ Missing = 'Integer' } -ErrorAction Stop } |
            Should -Throw '*not found in any DataSet worksheet*'
    }

    It 'appends rows without rewriting headers' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelAppend.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -AutoFit
        $moreRows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -Append -TableName 'Sales' -AutoFit

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B4')
        $imported.Count | Should -Be 3
        $imported[2].Region | Should -Be 'APAC'
        $imported[2].Revenue | Should -Be 150

        $excelSheetType = Get-TestLoadedType -Name 'OfficeIMO.Excel.ExcelSheet'
        $hasTableAppend = @($excelSheetType.GetMethods() | Where-Object Name -eq 'AppendDataTableToTable').Count -gt 0
        if ($hasTableAppend) {
            $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
            $tables.Count | Should -Be 1
            $tables[0].Range | Should -Be 'A1:B4'
        }
    }

    It 'can require appends to target an existing table' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelAppendToTable.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales'
        $moreRows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -Append -AppendToTable -TableName 'Sales'

        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A1:B4'

        $wrongRows = @([PSCustomObject]@{ Region = 'LATAM'; Amount = 400 })
        { $wrongRows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -Append -AppendToTable -TableName 'Sales' } |
            Should -Throw

        { $moreRows | Export-OfficeExcel -Path $path -WorksheetName 'Missing' -Append -AppendToTable -TableName 'Sales' } |
            Should -Throw
    }

    It 'exports DataTable input without exposing DataRow metadata' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataTable.xlsx'
        $table = [System.Data.DataTable]::new('Sales')
        [void] $table.Columns.Add('Region', [string])
        [void] $table.Columns.Add('Revenue', [int])
        [void] $table.Rows.Add('NA', 100)
        [void] $table.Rows.Add('EMEA', 200)

        Export-OfficeExcel -Path $path -InputObject $table -WorksheetName 'Data' -TableName 'Sales' -AutoFit

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[0].Region | Should -Be 'NA'
        $imported[0].Revenue | Should -Be 100
        $imported[0].PSObject.Properties.Name | Should -Not -Contain 'RowError'
    }

    It 'exports IDataReader input without requiring callers to buffer it first' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataReader.xlsx'
        $table = [System.Data.DataTable]::new('SqlRows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('A', 1)
        [void] $table.Rows.Add('B', 2)
        $reader = $table.CreateDataReader()

        Export-OfficeExcel -Path $path -InputObject $reader -WorksheetName 'Data' -TableName 'SqlRows' -AutoFit -FreezeTopRow

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data')
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'A'
        $rows[1].Value | Should -Be 2
    }

    It 'uses the compact package writer for an unmodified IDataReader export' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataReaderCompact.xlsx'
        $table = [System.Data.DataTable]::new('SqlRows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('A', 1)
        [void] $table.Rows.Add('B', 2)
        $reader = $table.CreateDataReader()

        Export-OfficeExcel -Path $path -InputObject $reader -WorksheetName 'Data' -TableName 'SqlRows'

        $reader.IsClosed | Should -BeFalse
        $archive = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $archive.GetEntry('xl/sharedStrings.xml') | Should -BeNullOrEmpty

            $sheetReader = [System.IO.StreamReader]::new($archive.GetEntry('xl/worksheets/sheet1.xml').Open())
            try {
                $sheetXml = $sheetReader.ReadToEnd()
            } finally {
                $sheetReader.Dispose()
            }

            $tableReader = [System.IO.StreamReader]::new($archive.GetEntry('xl/tables/table1.xml').Open())
            try {
                $tableXml = $tableReader.ReadToEnd()
            } finally {
                $tableReader.Dispose()
            }
        } finally {
            $archive.Dispose()
        }

        $sheetXml | Should -Match '<tableParts count="1">'
        $sheetXml | Should -Not -Match '<c r="A2"'
        $tableXml | Should -Match 'name="SqlRows"'
        $tableXml | Should -Match 'ref="A1:B3"'
        $tableXml | Should -Match '<autoFilter ref="A1:B3"'

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data')
        $rows.Count | Should -Be 2
        $rows[0].Name | Should -Be 'A'
        $rows[1].Value | Should -Be 2
    }

    It 'preserves an existing workbook when the compact IDataReader export fails' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataReaderFailure.xlsx'
        [PSCustomObject]@{ Name = 'Original'; Value = 42 } |
            Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'OriginalRows'
        $beforeHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash

        $table = [System.Data.DataTable]::new('SqlRows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Value', [int])
        [void] $table.Rows.Add('Replacement', 7)
        $reader = $table.CreateDataReader()
        $reader.Close()

        { Export-OfficeExcel -Path $path -InputObject $reader -WorksheetName 'Data' -TableName 'SqlRows' -ErrorAction Stop } |
            Should -Throw

        (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash | Should -Be $beforeHash
        @(Get-ChildItem -LiteralPath $TestDrive -Filter '.*.tmp.xlsx').Count |
            Should -Be 0
        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data')
        $rows.Count | Should -Be 1
        $rows[0].Name | Should -Be 'Original'
        $rows[0].Value | Should -Be 42
    }

    It 'honors NoClobber when the destination appears during compact IDataReader export' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataReaderNoClobberRace.xlsx'
        $table = [System.Data.DataTable]::new('SqlRows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Rows.Add('A')
        $innerReader = $table.CreateDataReader()
        $reader = [PSWriteOfficeTestSupport.LateDestinationDataReader]::new(
            $innerReader,
            $path,
            'late destination')

        { Export-OfficeExcel -Path $path -InputObject $reader -WorksheetName 'Data' -NoClobber -ErrorAction Stop } |
            Should -Throw

        [System.IO.File]::ReadAllText($path) | Should -Be 'late destination'
        $reader.IsClosed | Should -BeFalse
        @(Get-ChildItem -LiteralPath $TestDrive -Filter '.*.tmp.xlsx').Count |
            Should -Be 0
    }

    It 'rejects AppendToTable IDataReader export when the destination is missing' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataReaderAppendMissing.xlsx'
        $table = [System.Data.DataTable]::new('SqlRows')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Rows.Add('A')
        $reader = $table.CreateDataReader()

        { Export-OfficeExcel -Path $path -InputObject $reader -WorksheetName 'Data' -Append -AppendToTable -ErrorAction Stop } |
            Should -Throw

        $path | Should -Not -Exist
        $reader.IsClosed | Should -BeFalse
    }

    It 'exports HTML-parser DataTable output with companion link URL columns' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelHtmlDataTable.xlsx'
        $table = [System.Data.DataTable]::new('HtmlLinks')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('NameUrl', [string])
        [void] $table.Columns.Add('Status', [string])
        [void] $table.Rows.Add('Alpha', 'https://example.com/a', 'Ready')
        [void] $table.Rows.Add('Beta', 'https://example.com/b', 'Hold')

        $table | Export-OfficeExcel -Path $path -WorksheetName 'Links' -TableName 'HtmlLinks' -AutoFit -FreezeTopRow

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Links' -Range 'A1:C3')
        $imported.Count | Should -Be 2
        $imported[0].Name | Should -Be 'Alpha'
        $imported[0].NameUrl | Should -Be 'https://example.com/a'
        $imported[1].Status | Should -Be 'Hold'

        $tables = @(Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'HtmlLinks')
        $tables.Count | Should -Be 1
        $tables[0].Range | Should -Be 'A1:C3'
    }

    It 'exports DataSet input as one worksheet per table' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSet.xlsx'
        $dataSet = [System.Data.DataSet]::new('Report')

        $sales = [System.Data.DataTable]::new('Sales:2026')
        [void] $sales.Columns.Add('Region', [string])
        [void] $sales.Columns.Add('Revenue', [int])
        [void] $sales.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($sales)

        $inventory = [System.Data.DataTable]::new('Inventory')
        [void] $inventory.Columns.Add('Item', [string])
        [void] $inventory.Columns.Add('Count', [int])
        [void] $inventory.Rows.Add('Laptop', 5)
        [void] $dataSet.Tables.Add($inventory)

        Export-OfficeExcel -Path $path -InputObject $dataSet -TableName 'IgnoredForDataSet' -AutoFit

        $salesRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B2')
        $inventoryRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Inventory' -Range 'A1:B2')

        $salesRows.Count | Should -Be 1
        $salesRows[0].Region | Should -Be 'NA'
        $salesRows[0].Revenue | Should -Be 100
        $inventoryRows.Count | Should -Be 1
        $inventoryRows[0].Item | Should -Be 'Laptop'
        $inventoryRows[0].Count | Should -Be 5
    }

    It 'exports DataSet tables one at a time without mutating source tables' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetNamespaceDuplicates.xlsx'
        $dataSet = [System.Data.DataSet]::new('Report')
        $dataSet.Namespace = 'urn:report'

        $inheritedNamespace = [System.Data.DataTable]::new('T')
        [void] $inheritedNamespace.Columns.Add('Name', [string])
        [void] $inheritedNamespace.Columns.Add('Secret', [string])
        [void] $inheritedNamespace.Rows.Add('Inherited', 'one')
        [void] $dataSet.Tables.Add($inheritedNamespace)

        $emptyNamespace = [System.Data.DataTable]::new('T')
        $emptyNamespace.Namespace = ''
        [void] $emptyNamespace.Columns.Add('Name', [string])
        [void] $emptyNamespace.Columns.Add('Secret', [string])
        [void] $emptyNamespace.Rows.Add('Empty', 'two')
        [void] $dataSet.Tables.Add($emptyNamespace)

        Export-OfficeExcel -Path $path -InputObject $dataSet -ExcludeProperty Secret

        $firstRows = @(Import-OfficeExcel -Path $path -WorksheetName 'T' -Range 'A1:A2')
        $secondRows = @(Import-OfficeExcel -Path $path -WorksheetName 'T (2)' -Range 'A1:A2')

        $firstRows.Count | Should -Be 1
        $firstRows[0].Name | Should -Be 'Inherited'
        $secondRows.Count | Should -Be 1
        $secondRows[0].Name | Should -Be 'Empty'
        $inheritedNamespace.Columns.Contains('Secret') | Should -BeTrue
        $emptyNamespace.Columns.Contains('Secret') | Should -BeTrue
    }

    It 'appends and clears DataSet sheets using sanitized worksheet names' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetSanitizedAppend.xlsx'

        $dataSet = [System.Data.DataSet]::new('Report')
        $sales = [System.Data.DataTable]::new('Sales:2026')
        [void] $sales.Columns.Add('Region', [string])
        [void] $sales.Columns.Add('Revenue', [int])
        [void] $sales.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($sales)

        Export-OfficeExcel -Path $path -InputObject $dataSet

        $appendSet = [System.Data.DataSet]::new('Report')
        $appendSales = [System.Data.DataTable]::new('Sales:2026')
        [void] $appendSales.Columns.Add('Region', [string])
        [void] $appendSales.Columns.Add('Revenue', [int])
        [void] $appendSales.Rows.Add('EMEA', 200)
        [void] $appendSet.Tables.Add($appendSales)

        Export-OfficeExcel -Path $path -InputObject $appendSet -Append

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B3')
        $rows.Count | Should -Be 2
        $rows[1].Region | Should -Be 'EMEA'
        $rows[1].Revenue | Should -Be 200

        $replacementSet = [System.Data.DataSet]::new('Report')
        $replacementSales = [System.Data.DataTable]::new('Sales:2026')
        [void] $replacementSales.Columns.Add('Region', [string])
        [void] $replacementSales.Columns.Add('Revenue', [int])
        [void] $replacementSales.Rows.Add('APAC', 300)
        [void] $replacementSet.Tables.Add($replacementSales)

        Export-OfficeExcel -Path $path -InputObject $replacementSet -ClearSheet

        $replaced = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B2')
        $replaced.Count | Should -Be 1
        $replaced[0].Region | Should -Be 'APAC'
        $replaced[0].Revenue | Should -Be 300
    }

    It 'keeps sanitized symbol DataSet sheet names distinct from existing workbook sheets' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetFallbackCollision.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Sheet1' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Existing'
            }
        }

        $dataSet = [System.Data.DataSet]::new('Report')
        $table = [System.Data.DataTable]::new(':')
        [void] $table.Columns.Add('Region', [string])
        [void] $table.Columns.Add('Revenue', [int])
        [void] $table.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($table)

        Export-OfficeExcel -Path $path -InputObject $dataSet -Append

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $existingText = $null
            $doc['Sheet1'].TryGetCellText(1, 1, [ref] $existingText) | Should -BeTrue
            $existingText | Should -Be 'Existing'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName '_' -Range 'A1:B2')
        $rows.Count | Should -Be 1
        $rows[0].Region | Should -Be 'NA'
        $rows[0].Revenue | Should -Be 100

        $appendSet = [System.Data.DataSet]::new('Report')
        $appendTable = [System.Data.DataTable]::new(':')
        [void] $appendTable.Columns.Add('Region', [string])
        [void] $appendTable.Columns.Add('Revenue', [int])
        [void] $appendTable.Rows.Add('EMEA', 200)
        [void] $appendSet.Tables.Add($appendTable)

        Export-OfficeExcel -Path $path -InputObject $appendSet -Append

        $appendedRows = @(Import-OfficeExcel -Path $path -WorksheetName '_' -Range 'A1:B3')
        $appendedRows.Count | Should -Be 2
        $appendedRows[1].Region | Should -Be 'EMEA'
        $appendedRows[1].Revenue | Should -Be 200

        $replacementSet = [System.Data.DataSet]::new('Report')
        $replacementTable = [System.Data.DataTable]::new(':')
        [void] $replacementTable.Columns.Add('Region', [string])
        [void] $replacementTable.Columns.Add('Revenue', [int])
        [void] $replacementTable.Rows.Add('APAC', 300)
        [void] $replacementSet.Tables.Add($replacementTable)

        Export-OfficeExcel -Path $path -InputObject $replacementSet -ClearSheet

        $replacedRows = @(Import-OfficeExcel -Path $path -WorksheetName '_' -Range 'A1:B2')
        $replacedRows.Count | Should -Be 1
        $replacedRows[0].Region | Should -Be 'APAC'
        $replacedRows[0].Revenue | Should -Be 300

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $existingText = $null
            $doc['Sheet1'].TryGetCellText(1, 1, [ref] $existingText) | Should -BeTrue
            $existingText | Should -Be 'Existing'
            $doc.Sheets.Name | Should -Not -Contain '_ (2)'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'reuses existing suffixed DataSet sheets when appending and clearing sanitized duplicates' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetDuplicateSanitizedAppend.xlsx'

        $dataSet = [System.Data.DataSet]::new('Report')
        $salesColon = [System.Data.DataTable]::new('Sales:2026')
        [void] $salesColon.Columns.Add('Region', [string])
        [void] $salesColon.Columns.Add('Revenue', [int])
        [void] $salesColon.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($salesColon)

        $salesSlash = [System.Data.DataTable]::new('Sales/2026')
        [void] $salesSlash.Columns.Add('Region', [string])
        [void] $salesSlash.Columns.Add('Revenue', [int])
        [void] $salesSlash.Rows.Add('EMEA', 200)
        [void] $dataSet.Tables.Add($salesSlash)

        Export-OfficeExcel -Path $path -InputObject $dataSet

        $appendSet = [System.Data.DataSet]::new('Report')
        $appendColon = [System.Data.DataTable]::new('Sales:2026')
        [void] $appendColon.Columns.Add('Region', [string])
        [void] $appendColon.Columns.Add('Revenue', [int])
        [void] $appendColon.Rows.Add('APAC', 300)
        [void] $appendSet.Tables.Add($appendColon)

        $appendSlash = [System.Data.DataTable]::new('Sales/2026')
        [void] $appendSlash.Columns.Add('Region', [string])
        [void] $appendSlash.Columns.Add('Revenue', [int])
        [void] $appendSlash.Rows.Add('LATAM', 400)
        [void] $appendSet.Tables.Add($appendSlash)

        Export-OfficeExcel -Path $path -InputObject $appendSet -Append

        $firstRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B3')
        $secondRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026 (2)' -Range 'A1:B3')
        $firstRows.Count | Should -Be 2
        $secondRows.Count | Should -Be 2
        $firstRows[1].Region | Should -Be 'APAC'
        $secondRows[1].Region | Should -Be 'LATAM'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Name | Should -Not -Contain 'Sales_2026 (3)'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $replacementSet = [System.Data.DataSet]::new('Report')
        $replacementColon = [System.Data.DataTable]::new('Sales:2026')
        [void] $replacementColon.Columns.Add('Region', [string])
        [void] $replacementColon.Columns.Add('Revenue', [int])
        [void] $replacementColon.Rows.Add('NA', 500)
        [void] $replacementSet.Tables.Add($replacementColon)

        $replacementSlash = [System.Data.DataTable]::new('Sales/2026')
        [void] $replacementSlash.Columns.Add('Region', [string])
        [void] $replacementSlash.Columns.Add('Revenue', [int])
        [void] $replacementSlash.Rows.Add('EMEA', 600)
        [void] $replacementSet.Tables.Add($replacementSlash)

        Export-OfficeExcel -Path $path -InputObject $replacementSet -ClearSheet

        $replacedFirst = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026' -Range 'A1:B2')
        $replacedSecond = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales_2026 (2)' -Range 'A1:B2')
        $replacedFirst.Count | Should -Be 1
        $replacedSecond.Count | Should -Be 1
        $replacedFirst[0].Revenue | Should -Be 500
        $replacedSecond[0].Revenue | Should -Be 600
    }

    It 'preserves underscore-distinct DataSet sheet names and sparse suffix matches' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetSparseSuffixes.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Q1_Ops' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Existing'
            }
            Add-OfficeExcelSheet -Name 'Sparse' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
            }
            Add-OfficeExcelSheet -Name 'Sparse (10)' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
            }
        }

        $dataSet = [System.Data.DataSet]::new('Report')
        $underscored = [System.Data.DataTable]::new('Q1__Ops')
        [void] $underscored.Columns.Add('Region', [string])
        [void] $underscored.Columns.Add('Revenue', [int])
        [void] $underscored.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($underscored)

        $sparseFirst = [System.Data.DataTable]::new('Sparse ')
        [void] $sparseFirst.Columns.Add('Region', [string])
        [void] $sparseFirst.Columns.Add('Revenue', [int])
        [void] $sparseFirst.Rows.Add('EMEA', 200)
        [void] $dataSet.Tables.Add($sparseFirst)

        $sparseSecond = [System.Data.DataTable]::new(' Sparse')
        [void] $sparseSecond.Columns.Add('Region', [string])
        [void] $sparseSecond.Columns.Add('Revenue', [int])
        [void] $sparseSecond.Rows.Add('APAC', 300)
        [void] $dataSet.Tables.Add($sparseSecond)

        Export-OfficeExcel -Path $path -InputObject $dataSet -Append

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Name | Should -Contain 'Q1__Ops'
            $doc.Sheets.Name | Should -Not -Contain 'Q1_Ops (2)'
            $q1Text = $null
            $doc['Q1_Ops'].TryGetCellText(1, 1, [ref] $q1Text) | Should -BeTrue
            $q1Text | Should -Be 'Existing'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $underscoredRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Q1__Ops' -Range 'A1:B2')
        $firstRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sparse' -Range 'A1:B2')
        $sparseRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Sparse (10)' -Range 'A1:B2')

        $underscoredRows.Count | Should -Be 1
        $underscoredRows[0].Revenue | Should -Be 100
        $firstRows.Count | Should -Be 1
        $firstRows[0].Revenue | Should -Be 200
        $sparseRows.Count | Should -Be 1
        $sparseRows[0].Revenue | Should -Be 300
    }

    It 'matches the lowest suffixed DataSet sheet independent of workbook order' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetLowestSuffix.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data (3)' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
                Set-OfficeExcelCell -Address 'A2' -Value 'APAC'
                Set-OfficeExcelCell -Address 'B2' -Value 300
            }
            Add-OfficeExcelSheet -Name 'Data (2)' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'B1' -Value 'Revenue'
                Set-OfficeExcelCell -Address 'A2' -Value 'EMEA'
                Set-OfficeExcelCell -Address 'B2' -Value 200
            }
        }

        $dataSet = [System.Data.DataSet]::new('Report')
        $table = [System.Data.DataTable]::new('Data')
        [void] $table.Columns.Add('Region', [string])
        [void] $table.Columns.Add('Revenue', [int])
        [void] $table.Rows.Add('NA', 100)
        [void] $dataSet.Tables.Add($table)

        Export-OfficeExcel -Path $path -InputObject $dataSet -Append

        $lowestRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data (2)' -Range 'A1:B3')
        $higherRows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data (3)' -Range 'A1:B2')

        $lowestRows.Count | Should -Be 2
        $lowestRows[1].Region | Should -Be 'NA'
        $lowestRows[1].Revenue | Should -Be 100
        $higherRows.Count | Should -Be 1
        $higherRows[0].Revenue | Should -Be 300
    }

    It 'preserves symbol-only DataSet sheet names and sanitizes control characters' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelDataSetSymbolNames.xlsx'
        $controlName = "Bad$([char]1)Name"

        $dataSet = [System.Data.DataSet]::new('Report')
        foreach ($name in @('---', '___', $controlName)) {
            $table = [System.Data.DataTable]::new($name)
            [void] $table.Columns.Add('Region', [string])
            [void] $table.Columns.Add('Revenue', [int])
            [void] $table.Rows.Add('NA', 100)
            [void] $dataSet.Tables.Add($table)
        }

        Export-OfficeExcel -Path $path -InputObject $dataSet

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Name | Should -Contain '---'
            $doc.Sheets.Name | Should -Contain '___'
            $doc.Sheets.Name | Should -Contain 'Bad_Name'
            $doc.Sheets.Name | Should -Not -Contain 'Sheet1'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'adds a DataTable inside the Excel DSL table command' {
        $path = Join-Path $TestDrive 'DslExcelDataTable.xlsx'
        $table = [System.Data.DataTable]::new('Items')
        [void] $table.Columns.Add('Name', [string])
        [void] $table.Columns.Add('Quantity', [int])
        [void] $table.Rows.Add('Laptop', 5)
        [void] $table.Rows.Add('Tablet', 12)

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $table -TableName 'Items' -AutoFit
            }
        }

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B3')
        $imported.Count | Should -Be 2
        $imported[0].Name | Should -Be 'Laptop'
        $imported[0].Quantity | Should -Be 5
    }

    It 'lets OfficeIMO generate table names when the DSL caller omits them' {
        $path = Join-Path $TestDrive 'DslExcelGeneratedTableNames.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'First' -Content {
                Add-OfficeExcelTable -InputObject $rows
            }
            Add-OfficeExcelSheet -Name 'Second' -Content {
                Add-OfficeExcelTable -InputObject $rows
            }
        }

        $tables = @(Get-OfficeExcelTable -Path $path)
        $tables.Count | Should -Be 2
        @($tables.Name | Select-Object -Unique).Count | Should -Be 2
    }

    It 'keeps append freeze panes anchored to the existing table header' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelAppendFreeze.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        $rows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Sales' -Title 'Sales Export' -FreezeTopRow
        $moreRows | Export-OfficeExcel -Path $path -WorksheetName 'Data' -Append -TableName 'Sales' -FreezeTopRow

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $pane = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetViews']/*[local-name()='sheetView']/*[local-name()='pane']")

        $pane.GetAttribute('ySplit') | Should -Be '2'
    }

    It 'supports autofit and validation list helpers' {
        $path = Join-Path $TestDrive 'DslExcelExtras.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Status = 'New' }
            [PSCustomObject]@{ Name = 'Beta'; Status = 'Done' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Items' -AutoFit
                Add-OfficeExcelValidationList -Range 'C2:C3' -Values 'New','In Progress','Done'
                Invoke-OfficeExcelAutoFit -Columns
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'supports row/column helpers and reader metadata' {
        $path = Join-Path $TestDrive 'DslExcelReaders.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Value = 10 }
            [PSCustomObject]@{ Name = 'Beta'; Value = 20 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelColumn -Column 1 -StartRow 2 -Values 'Alpha', 'Beta'
                Set-OfficeExcelColumn -Column 2 -StartRow 2 -Values 10, 20
                Set-OfficeExcelNamedRange -Name 'ManualRange' -Range 'A1:B3'
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -StartRow 5
            }
        } | Out-Null

        $named = Get-OfficeExcelNamedRange -Path $path -Sheet 'Data' | Where-Object Name -eq 'ManualRange'
        $named | Should -Not -BeNullOrEmpty

        $namedRangeType = Get-TestLoadedType -Name 'PSWriteOffice.Cmdlets.Excel.GetOfficeExcelNamedRangeCommand'
        $normalizeRange = $namedRangeType.GetMethod('NormalizeRange', [System.Reflection.BindingFlags] 'NonPublic, Static')
        $normalizeRange.Invoke($null, @("'Budget`$2026'!`$A`$1:`$B`$2")) | Should -Be "'Budget`$2026'!A1:B2"

        $tables = Get-OfficeExcelTable -Path $path | Where-Object Name -eq 'Sales'
        $tables | Should -Not -BeNullOrEmpty

        $namedRows = @($named | Import-OfficeExcel)
        $namedRows.Count | Should -Be 2
        $namedRows[0].Name | Should -Be 'Alpha'
        $namedRows[0].Value | Should -Be 10

        $tableRows = @($tables | Import-OfficeExcel)
        $tableRows.Count | Should -Be 2
        $tableRows[1].Name | Should -Be 'Beta'
        $tableRows[1].Value | Should -Be 20

        $doc = Get-OfficeExcel -Path $path
        try {
            $documentRows = @($doc | Import-OfficeExcel -Sheet 'Data' -Range 'A1:B3')
            $documentRows.Count | Should -Be 2
            $documentRows[0].Name | Should -Be 'Alpha'

            $doc | Save-OfficeExcel | Out-Null
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $server = Start-TestHttpFileServer -FilePath $path -ContentType 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' -RequestCount 8
        try {
            $uri = [uri] $server.Url

            $remoteRows = @(Import-OfficeExcel -Uri $uri -AllowHttp -Sheet 'Data' -Range 'A1:B3')
            $remoteRows.Count | Should -Be 2
            $remoteRows[1].Name | Should -Be 'Beta'

            $remoteRange = @(Get-OfficeExcelRange -Uri $uri -AllowHttp -Sheet 'Data' -Range 'A1:B3')
            $remoteRange.Count | Should -Be 2
            $remoteRange[0].Value | Should -Be 10

            $remoteUsedRange = Get-OfficeExcelUsedRange -Uri $uri -AllowHttp -Sheet 'Data' -AsDataTable
            $remoteUsedRange.Rows.Count | Should -Be 6

            $remoteTables = Get-OfficeExcelTable -Uri $uri -AllowHttp | Where-Object Name -eq 'Sales'
            $remoteTableRows = @($remoteTables | Import-OfficeExcel -AllowHttp)
            $remoteTableRows.Count | Should -Be 2

            $remoteNamed = Get-OfficeExcelNamedRange -Uri $uri -AllowHttp -Sheet 'Data' | Where-Object Name -eq 'ManualRange'
            $remoteNamedRows = @($remoteNamed | Import-OfficeExcel -AllowHttp)
            $remoteNamedRows.Count | Should -Be 2

            $remoteDoc = Get-OfficeExcel -Uri $uri -AllowHttp -ReadOnly
            try {
                $remoteDocRows = @($remoteDoc | Import-OfficeExcel -Sheet 'Data' -Range 'A1:B3')
                $remoteDocRows.Count | Should -Be 2
                $remoteSummary = Get-OfficeExcelSummary -Document $remoteDoc -IncludeSheets -IncludeSchema
                $remoteSummary.SheetCount | Should -Be 1
                $remoteSummary.Sheets[0].Name | Should -Be 'Data'
                $remoteSummary.Schema.Worksheets[0].Name | Should -Be 'Data'
            } finally {
                Close-OfficeExcel -Document $remoteDoc
            }
        } finally {
            Stop-TestHttpFileServer -Server $server
        }
    }

    It 'applies gradient fills through the thin cell command' {
        $path = Join-Path $TestDrive 'DslExcelGradientFill.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Status' -GradientFrom '#FF0000' -GradientTo '#00FF00' -GradientDegree 45
            }
        }

        $stylesXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/styles.xml'
        $gradient = $stylesXml.SelectSingleNode("//*[local-name()='gradientFill']")
        $gradient | Should -Not -BeNullOrEmpty
        $gradient.GetAttribute('type') | Should -Be 'linear'
        $gradient.GetAttribute('degree') | Should -Be '45'

        $stops = @($gradient.SelectNodes("*[local-name()='stop']"))
        $stops.Count | Should -Be 2
        $stops[0].GetAttribute('position') | Should -Be '0'
        $stops[0].SelectSingleNode("*[local-name()='color']").GetAttribute('rgb') | Should -Be 'FFFF0000'
        $stops[1].GetAttribute('position') | Should -Be '1'
        $stops[1].SelectSingleNode("*[local-name()='color']").GetAttribute('rgb') | Should -Be 'FF00FF00'
    }

    It 'applies row layout through the thin row command' {
        $path = Join-Path $TestDrive 'DslExcelRowLayout.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Notes' -Bold $true -WrapText $true -Height 28 -FirstColumn 1 -LastColumn 2
                Set-OfficeExcelRow -Row 2 -Values 'Alpha', "Line 1`nLine 2" -AutoFit -Hidden $false
            }
        }

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $row1 = @($summary.Schema.Rows | Where-Object Index -eq 1)[0]
        $row1.CustomHeight | Should -BeTrue
        $row1.Height | Should -Be 28
        $summary.Schema.Worksheets[0].UsedRange | Should -Be 'A1:B2'
    }

    It 'applies worksheet outline groups through thin row and column commands' {
        $path = Join-Path $TestDrive 'DslExcelOutlineGroups.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Region', 'Q1', 'Q2', 'Total'
                Set-OfficeExcelRow -Row 2 -Values 'NA', 10, 12, 22
                Set-OfficeExcelRow -Row 3 -Values 'EMEA', 20, 30, 50
                Set-OfficeExcelRowGroup -StartRow 2 -EndRow 3 -Collapsed -SummaryBelow $true
                Set-OfficeExcelColumnGroup -StartColumn B -EndColumn C -Collapsed -SummaryRight $true
            }
        }

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'outlineLevel="1"'
        $worksheetXml | Should -Match 'hidden="1"'
        $worksheetXml | Should -Match 'collapsed="1"'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $summary.Schema.Worksheets[0].OutlineSummaryBelow | Should -BeTrue
        $summary.Schema.Worksheets[0].OutlineSummaryRight | Should -BeTrue

        $row2 = @($summary.Schema.Rows | Where-Object Index -eq 2)[0]
        $row4 = @($summary.Schema.Rows | Where-Object Index -eq 4)[0]
        $row2.OutlineLevel | Should -Be 1
        $row2.Hidden | Should -BeTrue
        $row4.Collapsed | Should -BeTrue

        $column2 = @($summary.Schema.Columns | Where-Object StartIndex -eq 2)[0]
        $column4 = @($summary.Schema.Columns | Where-Object StartIndex -eq 4)[0]
        $column2.OutlineLevel | Should -Be 1
        $column2.Hidden | Should -BeTrue
        $column4.Collapsed | Should -BeTrue
    }

    It 'adds subtotal summaries through a thin worksheet workflow command' {
        $path = Join-Path $TestDrive 'DslExcelSubtotalSummary.xlsx'
        $script:subtotal = $null

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Region', 'Sales', 'Units'
                Set-OfficeExcelRow -Row 2 -Values 'NA', 100, 2
                Set-OfficeExcelRow -Row 3 -Values 'NA', 150, 3
                Set-OfficeExcelRow -Row 4 -Values 'EMEA', 200, 4
                Set-OfficeExcelRow -Row 5 -Values 'EMEA', 50, 1

                $script:subtotal = Add-OfficeExcelSubtotalSummary -GroupColumn Region -ValueColumn Sales, Units -DataEndRow 5 -SummaryStartRow 7 -PassThru
            }
        }

        $script:subtotal.SummaryRange | Should -Be 'A7:C10'
        $script:subtotal.GroupCount | Should -Be 2
        $script:subtotal.GrandTotalWritten | Should -BeTrue
        Get-TestWorksheetCellFormula -Path $path -Address B8 | Should -Be 'SUBTOTAL(9,B2:B3)'
        Get-TestWorksheetCellFormula -Path $path -Address C10 | Should -Be 'SUBTOTAL(9,C2:C5)'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        (@($summary.Schema.Rows | Where-Object Index -eq 2)[0]).OutlineLevel | Should -Be 1
        (@($summary.Schema.Rows | Where-Object Index -eq 4)[0]).OutlineLevel | Should -Be 1

        $shortHeaderPath = Join-Path $TestDrive 'DslExcelSubtotalSummaryShortHeaders.xlsx'
        New-OfficeExcel -Path $shortHeaderPath {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'ID', 'SKU', 'Amount'
                Set-OfficeExcelRow -Row 2 -Values 'A', 10, 2
                Set-OfficeExcelRow -Row 3 -Values 'A', 15, 3
                Add-OfficeExcelSubtotalSummary -GroupColumn ID -ValueColumn SKU -DataEndRow 3 -SummaryStartRow 5
            }
        }

        Get-TestWorksheetCellFormula -Path $shortHeaderPath -Address B6 | Should -Be 'SUBTOTAL(9,B2:B3)'
    }

    It 'sets and clears worksheet tab colors through thin sheet commands' {
        $path = Join-Path $TestDrive 'DslExcelSheetTabColor.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelSheetTabColor -Color '#336699'
            }
        }

        $worksheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $tabColor = $worksheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='sheetPr']/*[local-name()='tabColor']")
        $tabColor.GetAttribute('rgb') | Should -Be 'FF336699'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $summary.Schema.Worksheets[0].TabColorArgb | Should -Be 'FF336699'

        $doc = Get-OfficeExcel -Path $path
        try {
            $doc | Set-OfficeExcelSheetTabColor -Sheet Data -Clear
        } finally {
            Close-OfficeExcel -Document $doc -Save
        }

        $summaryAfterClear = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $summaryAfterClear.Schema.Worksheets[0].TabColorArgb | Should -BeNullOrEmpty
    }

    It 'sets worksheet view options through thin sheet commands' {
        $path = Join-Path $TestDrive 'DslExcelWorksheetViewOptions.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelWorksheetView -HideGridlines -RightToLeft -ZoomScale 125 -ZoomScaleNormal 100 -View PageLayout
            }
        }

        $view = Get-OfficeExcelWorksheetView -Path $path -Sheet Data
        $view.ShowGridlines | Should -BeFalse
        $view.RightToLeft | Should -BeTrue
        $view.ZoomScale | Should -Be 125
        $view.ZoomScaleNormal | Should -Be 100
        $view.View | Should -Be 'pageLayout'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSchema
        $summary.Schema.Worksheets[0].ShowGridlines | Should -BeFalse
        $summary.Schema.Worksheets[0].RightToLeft | Should -BeTrue
        $summary.Schema.Worksheets[0].ZoomScale | Should -Be 125
        $summary.Schema.Worksheets[0].View | Should -Be 'pageLayout'

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'showGridLines="(?:0|false)"'
        $worksheetXml | Should -Match 'rightToLeft="(?:1|true)"'
        $worksheetXml | Should -Match 'zoomScale="125"'
        $worksheetXml | Should -Match 'view="pageLayout"'

        $doc = Get-OfficeExcel -Path $path
        try {
            $doc | Set-OfficeExcelWorksheetView -Sheet Data -ShowGridlines -LeftToRight -ZoomScale 90 -View Normal
        } finally {
            Close-OfficeExcel -Document $doc -Save
        }

        $viewAfterUpdate = Get-OfficeExcelWorksheetView -Path $path -Sheet Data
        $viewAfterUpdate.ShowGridlines | Should -BeTrue
        $viewAfterUpdate.RightToLeft | Should -BeFalse
        $viewAfterUpdate.ZoomScale | Should -Be 90
        $viewAfterUpdate.View | Should -Be 'normal'
    }

    It 'reads worksheet view options from the current DSL sheet' {
        $path = Join-Path $TestDrive 'DslExcelWorksheetViewContext.xlsx'
        $script:currentSheetView = $null

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelWorksheetView -HideGridlines -ZoomScale 125 -View PageLayout
                $script:currentSheetView = Get-OfficeExcelWorksheetView
            }
        }

        $script:currentSheetView.SheetName | Should -Be 'Data'
        $script:currentSheetView.ShowGridlines | Should -BeFalse
        $script:currentSheetView.ZoomScale | Should -Be 125
        $script:currentSheetView.View | Should -Be 'pageLayout'
    }

    It 'sets the active worksheet through thin workbook commands' {
        $path = Join-Path $TestDrive 'DslExcelActiveSheet.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Summary'
            }
            Add-OfficeExcelSheet -Name 'Details' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Details'
                Set-OfficeExcelActiveSheet
            }
            Add-OfficeExcelSheet -Name 'Archive' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Archive'
            }
        }

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets -IncludeSchema
        $summary.ActiveSheetIndex | Should -Be 1
        $summary.ActiveSheetName | Should -Be 'Details'
        (@($summary.Sheets | Where-Object IsActive).Name) | Should -Be 'Details'
        $summary.Schema.ActiveWorksheetIndex | Should -Be 1
        $summary.Schema.ActiveWorksheetName | Should -Be 'Details'
        (@($summary.Schema.Worksheets | Where-Object IsActive).Name) | Should -Be 'Details'

        $workbookXml = Read-XlsxEntryText -Path $path -Entry 'xl/workbook.xml'
        $workbookXml | Should -Match 'activeTab="1"'
        $detailsXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet2.xml'
        $detailsXml | Should -Match 'tabSelected="(?:1|true)"'

        $activeResult = Set-OfficeExcelActiveSheet -Path $path -SheetIndex 2 -PassThru
        $activeResult.Path | Should -Be $path
        $activeResult.Name | Should -Be 'Archive'
        $activeResult.SheetName | Should -Be 'Archive'
        $activeResult.SheetIndex | Should -Be 2

        $updatedSummary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $updatedSummary.ActiveSheetIndex | Should -Be 2
        $updatedSummary.ActiveSheetName | Should -Be 'Archive'
        (@($updatedSummary.Sheets | Where-Object IsActive).Name) | Should -Be 'Archive'
    }

    It 'imports formula text while accepting explicit culture options' {
        $path = Join-Path $TestDrive 'DslExcelFormulaImport.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Amount', 'Formula'
                Set-OfficeExcelRow -Row 2 -Values 'Alpha', '1,25', ''
                Set-OfficeExcelFormula -Address 'C2' -Formula 'SUM(1,2)'
            }
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:C2' -FormulaMode FormulaText -CultureName 'pl-PL')
        $rows[0].Amount | Should -Be '1,25'
        $rows[0].Formula | Should -Be 'SUM(1,2)'
    }

    It 'saves evaluated formula caches through friendly save options' {
        $path = Join-Path $TestDrive 'DslExcelFormulaSaveOptions.xlsx'

        New-OfficeExcel -Path $path -EvaluateFormulas {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'A', 'B', 'Total'
                Set-OfficeExcelRow -Row 2 -Values 10, 15, ''
                Set-OfficeExcelFormula -Address 'C2' -Formula 'SUM(A2:B2)'
            }
        }

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:C2')
        $rows[0].Total | Should -Be 25
    }

    It 'sets and reads Excel document metadata' {
        $path = Join-Path $TestDrive 'DslExcelMetadata.xlsx'

        New-OfficeExcel -Path $path -DocumentTitle 'Operations Report' -Author 'PSWriteOffice' -Company 'Evotec' {
            Set-OfficeExcelDocumentProperty -Name Subject -Value 'Spreadsheet metadata'
            Set-OfficeExcelDocumentProperty -Name ReleaseStatus -Value Approved -Custom
            Set-OfficeExcelDocumentProperty -Name Ticket -Value 42 -Custom
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
            }
        }

        $properties = @(Get-OfficeExcelDocumentProperty -Path $path -Name Title, Creator, Subject, Company, ReleaseStatus, Ticket)
        ($properties | Where-Object Name -eq Title).Value | Should -Be 'Operations Report'
        ($properties | Where-Object Name -eq Creator).Value | Should -Be 'PSWriteOffice'
        ($properties | Where-Object Name -eq Subject).Value | Should -Be 'Spreadsheet metadata'
        ($properties | Where-Object Name -eq Company).Value | Should -Be 'Evotec'
        ($properties | Where-Object Name -eq ReleaseStatus).Value | Should -Be 'Approved'
        ($properties | Where-Object Name -eq Ticket).Value | Should -Be 42
        ($properties | Where-Object Name -eq Ticket).Scope | Should -Be 'Custom'
    }

    It 'applies Excel template markers through a thin template command' {
        $path = Join-Path $TestDrive 'DslExcelTemplate.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Invoice' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Invoice {{Number}}'
                Set-OfficeExcelCell -Address A2 -Value 'Total {{Total:currency}}'
                Set-OfficeExcelCell -Address A3 -Value 'Missing {{Optional}}'
            }
        }

        Invoke-OfficeExcelTemplate -Path $path -Sheet Invoice -Value @{ Number = 'INV-001'; Total = 123.4 } -CultureName en-US -MissingValueBehavior EmptyString -PassThru | Should -Be 3

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'Invoice INV-001'
        $worksheetXml | Should -Match 'Total \$123\.40'
        $worksheetXml | Should -Match 'Missing '
        $worksheetXml | Should -Not -Match '\{\{'
    }

    It 'does not save path-owned template work when WhatIf skips all sheets' {
        $path = Join-Path $TestDrive 'DslExcelTemplateWhatIf.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Invoice' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Invoice {{Number}}'
            }
        }

        $before = [System.IO.File]::ReadAllBytes($path)
        Invoke-OfficeExcelTemplate -Path $path -Sheet Invoice -Value @{ Number = 'INV-001' } -WhatIf
        $after = [System.IO.File]::ReadAllBytes($path)

        [Convert]::ToBase64String($after) | Should -Be ([Convert]::ToBase64String($before))
    }

    It 'inspects Excel template markers and reports missing bindings' {
        $path = Join-Path $TestDrive 'DslExcelTemplateInspection.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Invoice' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Invoice {{Number}}'
                Set-OfficeExcelCell -Address A2 -Value 'Total {{Total:currency}}'
                Set-OfficeExcelCell -Address A3 -Value 'Missing {{Optional}}'
                $script:contextTemplateMarkers = @(Get-OfficeExcelTemplateMarker -Value @{ Number = 'INV-001'; Total = 123.4 })
            }
        }

        $script:contextTemplateMarkers.Count | Should -Be 3
        ($script:contextTemplateMarkers | Where-Object Name -eq Number).IsBound | Should -BeTrue

        $markers = @(Get-OfficeExcelTemplateMarker -Path $path -Sheet Invoice -Value @{ Number = 'INV-001'; Total = 123.4 })
        $markers.Count | Should -Be 3
        ($markers | Where-Object Name -eq Number).Address | Should -Be 'A1'
        ($markers | Where-Object Name -eq Total).Format | Should -Be 'currency'
        ($markers | Where-Object Name -eq Total).IsBound | Should -BeTrue

        $missing = @(Get-OfficeExcelTemplateMarker -Path $path -Sheet Invoice -Value @{ Number = 'INV-001'; Total = 123.4 } -MissingOnly)
        $missing.Count | Should -Be 1
        $missing[0].Name | Should -Be 'Optional'
        $missing[0].IsBound | Should -BeFalse
    }

    It 'repeats an Excel template row from pipeline data' {
        $path = Join-Path $TestDrive 'DslExcelTemplateRows.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Invoice' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Item'
                Set-OfficeExcelCell -Address B1 -Value 'Amount'
                Set-OfficeExcelCell -Address A2 -Value '{{Name}}'
                Set-OfficeExcelCell -Address B2 -Value 'Amount {{Amount:currency}}'
                Set-OfficeExcelCell -Address A3 -Value 'Footer'
            }
        }

        @(
            [pscustomobject]@{ Name = 'Consulting'; Amount = 1200 }
            [pscustomobject]@{ Name = 'Support'; Amount = 300 }
        ) | Invoke-OfficeExcelTemplateRow -Path $path -Sheet Invoice -TemplateRow 2 -CultureName en-US -MissingValueBehavior Throw -PassThru | Should -Be 4

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'Consulting'
        $worksheetXml | Should -Match 'Support'
        $worksheetXml | Should -Match 'Amount \$1,200\.00'
        $worksheetXml | Should -Match 'Amount \$300\.00'
        $worksheetXml | Should -Match 'r="A4"'
        $worksheetXml | Should -Not -Match '\{\{'
        (Read-XlsxEntryText -Path $path -Entry 'xl/sharedStrings.xml') | Should -Match 'Footer'
    }

    It 'repeats an Excel template row from generic dictionary data' {
        $path = Join-Path $TestDrive 'DslExcelTemplateRows.GenericDictionary.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Invoice' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Item'
                Set-OfficeExcelCell -Address B1 -Value 'Amount'
                Set-OfficeExcelCell -Address A2 -Value '{{Name}}'
                Set-OfficeExcelCell -Address B2 -Value 'Amount {{Amount:currency}}'
                Set-OfficeExcelCell -Address A3 -Value 'Footer'
            }
        }

        $first = [System.Collections.Generic.Dictionary[string,object]]::new()
        $first['Name'] = 'Consulting'
        $first['Amount'] = 1200
        Invoke-OfficeExcelTemplateRow -Path $path -Sheet Invoice -TemplateRow 2 -InputObject $first -CultureName en-US -MissingValueBehavior Throw -PassThru | Should -Be 2

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'Consulting'
        $worksheetXml | Should -Match 'Amount \$1,200\.00'
        $worksheetXml | Should -Not -Match '\{\{'

        $readOnlyPath = Join-Path $TestDrive 'DslExcelTemplateRows.ReadOnlyDictionary.xlsx'
        New-OfficeExcel -Path $readOnlyPath {
            Add-OfficeExcelSheet -Name 'Invoice' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Item'
                Set-OfficeExcelCell -Address B1 -Value 'Amount'
                Set-OfficeExcelCell -Address A2 -Value '{{Name}}'
                Set-OfficeExcelCell -Address B2 -Value 'Amount {{Amount:currency}}'
            }
        }

        $backing = [System.Collections.Generic.Dictionary[string,object]]::new()
        $backing['Name'] = 'Support'
        $backing['Amount'] = 300
        $readOnly = [System.Collections.ObjectModel.ReadOnlyDictionary[string,object]]::new($backing)

        Invoke-OfficeExcelTemplateRow -Path $readOnlyPath -Sheet Invoice -TemplateRow 2 -Rows $readOnly -CultureName en-US -MissingValueBehavior Throw -PassThru | Should -Be 2

        $worksheetXml = Read-XlsxEntryText -Path $readOnlyPath -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'Support'
        $worksheetXml | Should -Match 'Amount \$300\.00'
        $worksheetXml | Should -Not -Match '\{\{'
    }

    It 'expands named Excel template row arrays' {
        $path = Join-Path $TestDrive 'DslExcelTemplateRows.NamedArray.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Invoice' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Item'
                Set-OfficeExcelCell -Address B1 -Value 'Amount'
                Set-OfficeExcelCell -Address A2 -Value '{{Name}}'
                Set-OfficeExcelCell -Address B2 -Value '{{Amount}}'
            }
        }

        $items = @(
            [pscustomobject]@{ Name = 'Consulting'; Amount = 1200 }
            [pscustomobject]@{ Name = 'Support'; Amount = 300 }
        )

        Invoke-OfficeExcelTemplateRow -Path $path -Sheet Invoice -TemplateRow 2 -Rows $items -MissingValueBehavior Throw -PassThru | Should -Be 4

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'Consulting'
        $worksheetXml | Should -Match 'Support'
        $worksheetXml | Should -Match 'r="A3"'
        $worksheetXml | Should -Not -Match '\{\{'
    }

    It 'includes and removes optional Excel template rows' {
        $path = Join-Path $TestDrive 'DslExcelTemplateOptionalRows.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Invoice' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Header'
                Set-OfficeExcelCell -Address A2 -Value 'Discount {{Discount}}'
                Set-OfficeExcelCell -Address A3 -Value 'Footer'
            }
        }

        Invoke-OfficeExcelTemplateOptionalRow -Path $path -Sheet Invoice -FirstRow 2 -Value @{ Discount = '10%' } -MissingValueBehavior Throw -PassThru | Should -Be 1
        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'Discount 10%'
        $worksheetXml | Should -Not -Match '\{\{'

        Invoke-OfficeExcelTemplateOptionalRow -Path $path -Sheet Invoice -FirstRow 2 -RowCount 1 -Remove
        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'ref="A1:A2"'
        $worksheetXml | Should -Match 'r="A2"'
        $worksheetXml | Should -Not -Match 'Discount 10%'
        (Read-XlsxEntryText -Path $path -Entry 'xl/sharedStrings.xml') | Should -Match 'Footer'
    }

    It 'repeats Excel template sheets from pipeline data' {
        $path = Join-Path $TestDrive 'DslExcelTemplateSheets.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Template' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Invoice {{Number}}'
                Set-OfficeExcelCell -Address A2 -Value 'Customer {{Customer}}'
            }
        }

        @(
            [pscustomobject]@{ Number = 'INV-001'; Customer = 'Contoso'; SheetName = 'Inv-001' }
            [pscustomobject]@{ Number = 'INV-002'; Customer = 'Fabrikam'; SheetName = 'Inv-002' }
        ) | Invoke-OfficeExcelTemplateSheet -Path $path -TemplateSheet Template -SheetNameProperty SheetName -MissingValueBehavior Throw -PassThru | Should -Be 4

        $workbookXml = Read-XlsxEntryText -Path $path -Entry 'xl/workbook.xml'
        $workbookXml | Should -Match 'name="Inv-001"'
        $workbookXml | Should -Match 'name="Inv-002"'

        $worksheetEntries = @(Get-ZipEntriesLocal -Path $path | Where-Object { $_ -like 'xl/worksheets/sheet*.xml' })
        $worksheetEntries.Count | Should -Be 2
        $worksheetsXml = ($worksheetEntries | ForEach-Object { Read-XlsxEntryText -Path $path -Entry $_ }) -join [Environment]::NewLine
        $worksheetsXml | Should -Match 'Invoice INV-001'
        $worksheetsXml | Should -Match 'Customer Contoso'
        $worksheetsXml | Should -Match 'Invoice INV-002'
        $worksheetsXml | Should -Match 'Customer Fabrikam'
        $worksheetsXml | Should -Not -Match '\{\{'
    }

    It 'expands named Excel template sheet arrays' {
        $path = Join-Path $TestDrive 'DslExcelTemplateSheets.NamedArray.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Template' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Invoice {{Number}}'
                Set-OfficeExcelCell -Address A2 -Value 'Customer {{Customer}}'
            }
        }

        $invoices = @(
            [pscustomobject]@{ Number = 'INV-101'; Customer = 'Northwind'; SheetName = 'Inv-101' }
            [pscustomobject]@{ Number = 'INV-102'; Customer = 'Adventure Works'; SheetName = 'Inv-102' }
        )

        Invoke-OfficeExcelTemplateSheet -Path $path -TemplateSheet Template -SheetNameProperty SheetName -Rows $invoices -MissingValueBehavior Throw -PassThru | Should -Be 4

        $workbookXml = Read-XlsxEntryText -Path $path -Entry 'xl/workbook.xml'
        $workbookXml | Should -Match 'name="Inv-101"'
        $workbookXml | Should -Match 'name="Inv-102"'

        $worksheetEntries = @(Get-ZipEntriesLocal -Path $path | Where-Object { $_ -like 'xl/worksheets/sheet*.xml' })
        $worksheetEntries.Count | Should -Be 2
        $worksheetsXml = ($worksheetEntries | ForEach-Object { Read-XlsxEntryText -Path $path -Entry $_ }) -join [Environment]::NewLine
        $worksheetsXml | Should -Match 'Invoice INV-101'
        $worksheetsXml | Should -Match 'Customer Northwind'
        $worksheetsXml | Should -Match 'Invoice INV-102'
        $worksheetsXml | Should -Match 'Customer Adventure Works'
        $worksheetsXml | Should -Not -Match '\{\{'

        $readOnlyPath = Join-Path $TestDrive 'DslExcelTemplateSheets.ReadOnlyDictionary.xlsx'
        New-OfficeExcel -Path $readOnlyPath {
            Add-OfficeExcelSheet -Name 'Template' -Content {
                Set-OfficeExcelCell -Address A1 -Value 'Invoice {{Number}}'
                Set-OfficeExcelCell -Address A2 -Value 'Customer {{Customer}}'
            }
        }

        $backing = [System.Collections.Generic.Dictionary[string,object]]::new()
        $backing['Number'] = 'INV-201'
        $backing['Customer'] = 'Litware'
        $backing['SheetName'] = 'Inv-201'
        $readOnly = [System.Collections.ObjectModel.ReadOnlyDictionary[string,object]]::new($backing)

        Invoke-OfficeExcelTemplateSheet -Path $readOnlyPath -TemplateSheet Template -SheetNameProperty SheetName -Rows $readOnly -MissingValueBehavior Throw -PassThru | Should -Be 2

        $workbookXml = Read-XlsxEntryText -Path $readOnlyPath -Entry 'xl/workbook.xml'
        $workbookXml | Should -Match 'name="Inv-201"'
        $worksheetsXml = (Get-ZipEntriesLocal -Path $readOnlyPath | Where-Object { $_ -like 'xl/worksheets/sheet*.xml' } | ForEach-Object { Read-XlsxEntryText -Path $readOnlyPath -Entry $_ }) -join [Environment]::NewLine
        $worksheetsXml | Should -Match 'Invoice INV-201'
        $worksheetsXml | Should -Match 'Customer Litware'
        $worksheetsXml | Should -Not -Match '\{\{'
    }

    It 'sets named ranges, formulas, and header/footer' {
        $path = Join-Path $TestDrive 'DslExcelMeta.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 10
                Set-OfficeExcelCell -Address 'B1' -Value 20
                Set-OfficeExcelFormula -Address 'C1' -Formula 'SUM(A1:B1)'
                Set-OfficeExcelNamedRange -Name 'Totals' -Range 'A1:C1'
                Set-OfficeExcelHeaderFooter -HeaderCenter 'Demo' -FooterRight 'Page &P of &N'
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'applies friendly AutoFilter conditions by header' {
        $path = Join-Path $TestDrive 'DslExcelAutoFilterHeaders.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Region', 'Sales', 'Notes'
                Set-OfficeExcelRow -Row 2 -Values 'NA', 100, 'urgent review'
                Set-OfficeExcelRow -Row 3 -Values 'EMEA', 200, 'normal'
                Set-OfficeExcelRow -Row 4 -Values 'APAC', 300, 'urgent hold'
                Set-OfficeExcelAutoFilter -Range 'A1:C4' -Header Region -Value NA, EMEA
                Set-OfficeExcelAutoFilter -Header Sales -Between 100, 250
                Set-OfficeExcelAutoFilter -Header Notes -Contains urgent
            }
        }

        $worksheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $autoFilter = $worksheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='autoFilter']")
        $autoFilter.GetAttribute('ref') | Should -Be 'A1:C4'

        $regionFilters = @($autoFilter.SelectNodes("*[local-name()='filterColumn'][@colId='0']/*[local-name()='filters']/*[local-name()='filter']") | ForEach-Object { $_.GetAttribute('val') })
        $regionFilters | Should -Be @('NA', 'EMEA')

        $salesFilters = @($autoFilter.SelectNodes("*[local-name()='filterColumn'][@colId='1']/*[local-name()='customFilters']/*[local-name()='customFilter']"))
        $salesFilters.Count | Should -Be 2
        $salesFilters[0].GetAttribute('operator') | Should -Be 'greaterThanOrEqual'
        $salesFilters[0].GetAttribute('val') | Should -Be '100'
        $salesFilters[1].GetAttribute('operator') | Should -Be 'lessThanOrEqual'
        $salesFilters[1].GetAttribute('val') | Should -Be '250'

        $notesFilter = $autoFilter.SelectSingleNode("*[local-name()='filterColumn'][@colId='2']/*[local-name()='customFilters']/*[local-name()='customFilter']")
        $notesFilter.GetAttribute('operator') | Should -Be 'equal'
        $notesFilter.GetAttribute('val') | Should -Be '*urgent*'
    }

    It 'supports advanced Excel data helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedData.xlsx'
        $rows = @(
            [PSCustomObject]@{
                Region = 'NA'
                Sales = 100
                Rate = 0.2
                CloseDate = [datetime]'2024-01-15'
                StartTime = [TimeSpan]'08:30:00'
                Note = 'OK'
            }
            [PSCustomObject]@{
                Region = 'EMEA'
                Sales = 200
                Rate = 0.45
                CloseDate = [datetime]'2024-02-20'
                StartTime = [TimeSpan]'09:15:00'
                Note = 'Check'
            }
            [PSCustomObject]@{
                Region = 'APAC'
                Sales = 150
                Rate = 0.33
                CloseDate = [datetime]'2024-03-10'
                StartTime = [TimeSpan]'10:05:00'
                Note = 'Review'
            }
        )

        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                Add-OfficeExcelAutoFilter -Range 'A1:F4'
                Invoke-OfficeExcelSort -Header 'Region'
                Set-OfficeExcelFreeze -TopRows 1
                Add-OfficeExcelConditionalRule -Range 'B2:B4' -Operator GreaterThan -Formula1 '150'
                Add-OfficeExcelConditionalDataBar -Range 'B2:B4' -Color '#FF4F81BD'
                Add-OfficeExcelConditionalColorScale -Range 'C2:C4' -StartColor '#FE9' -EndColor '#6A84'
                Add-OfficeExcelConditionalIconSet -Range 'C2:C4'
                Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type ColumnClustered -Title 'Sales'
                Add-OfficeExcelImage -Address 'I1' -Path $imagePath -WidthPixels 64 -HeightPixels 64
                Set-OfficeExcelHyperlink -Address 'A2' -Url 'https://example.org' -Display 'Example'
                Add-OfficeExcelComment -Address 'B2' -Text 'Check sales' -Author Alice -Initials AA
                $script:contextComments = @(Get-OfficeExcelComment -TextContains sales)
                Add-OfficeExcelSparkline -DataRange 'B2:B4' -LocationRange 'H2:H4' -Type Column
            }
        }

        Test-Path $path | Should -BeTrue

        $conditionalXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $conditionalXml | Should -Match 'rgb="FF4F81BD"'
        $conditionalXml | Should -Match 'rgb="FFFFEE99"'
        $conditionalXml | Should -Match 'rgb="4466AA88"'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -Be 1
            $sheet = $doc.Sheets[0]
            $sheet.Name | Should -Be 'Data'
            $sheet.HasComment(2, 2) | Should -BeTrue

            $cellText = $null
            $sheet.TryGetCellText(2, 1, [ref] $cellText) | Should -BeTrue
            $cellText | Should -Be 'Example'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $script:contextComments.Count | Should -Be 1
        $script:contextComments[0].SheetName | Should -Be 'Data'

        $comments = @(Get-OfficeExcelComment -Path $path -Sheet Data -TextContains sales)
        $comments.Count | Should -Be 1
        $comments[0].Address | Should -Be 'B2'
        $comments[0].Author | Should -Be 'Alice (AA)'
        $comments[0].Text | Should -Be 'Check sales'

        Update-OfficeExcelComment -Path $path -Sheet Data -Address B2 -Text 'Should not save' -Author Nope -WhatIf
        $whatIfComment = Get-OfficeExcelComment -Path $path -Sheet Data -Address B2
        $whatIfComment.Author | Should -Be 'Alice (AA)'
        $whatIfComment.Text | Should -Be 'Check sales'

        Update-OfficeExcelComment -Path $path -Sheet Data -Address B2 -Text 'Reviewed sales' -Author Carol -Initials CC -PassThru | Should -Be 1
        $updatedComment = Get-OfficeExcelComment -Path $path -Sheet Data -Address B2
        $updatedComment.Author | Should -Be 'Carol (CC)'
        $updatedComment.Text | Should -Be 'Reviewed sales'

        Clear-OfficeExcelComment -Path $path -Sheet Data -TextContains Reviewed -PassThru -Confirm:$false | Should -Be 1
        @(Get-OfficeExcelComment -Path $path -Sheet Data -Address B2).Count | Should -Be 0
    }

    It 'supports advanced Excel pivot, validation, and protection helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedPivot.xlsx'
        $rows = @(
            [PSCustomObject]@{
                Region = 'NA'
                Sales = 100
                Rate = 0.2
                CloseDate = [datetime]'2024-01-15'
                StartTime = [TimeSpan]'08:30:00'
                Note = 'OK'
            }
            [PSCustomObject]@{
                Region = 'EMEA'
                Sales = 200
                Rate = 0.45
                CloseDate = [datetime]'2024-02-20'
                StartTime = [TimeSpan]'09:15:00'
                Note = 'Check'
            }
            [PSCustomObject]@{
                Region = 'APAC'
                Sales = 150
                Rate = 0.33
                CloseDate = [datetime]'2024-03-10'
                StartTime = [TimeSpan]'10:05:00'
                Note = 'Review'
            }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                Add-OfficeExcelPivotTable -SourceRange 'A1:F4' -DestinationCell 'J1' -RowField 'Region' -DataField 'Sales' -DataDisplayName 'Total Sales'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B4' -Operator Between -Formula1 1 -Formula2 1000 -AllowBlank:$false
                Add-OfficeExcelValidationDecimal -Range 'C2:C4' -Operator Between -Formula1 0.0 -Formula2 1.0
                Add-OfficeExcelValidationDate -Range 'D2:D4' -Operator GreaterThan -Formula1 ([datetime]'2024-01-01')
                Add-OfficeExcelValidationTime -Range 'E2:E4' -Operator GreaterThan -Formula1 ([TimeSpan]'08:00:00')
                Add-OfficeExcelValidationTextLength -Range 'F2:F4' -Operator Between -Formula1 1 -Formula2 20
                Add-OfficeExcelValidationCustomFormula -Range 'G2:G4' -Formula 'LEN(A2)>0'
                Protect-OfficeExcelSheet
                Unprotect-OfficeExcelSheet
                Protect-OfficeExcelSheet
            }
        }

        Test-Path $path | Should -BeTrue

        $pivotTables = @(Get-OfficeExcelPivotTable -Path $path -Name 'PivotTable')
        $pivotTables.Count | Should -Be 1

        $pivot = $pivotTables[0]
        $pivot.SourceRange | Should -Be 'A1:F4'
        $pivot.Location | Should -Match '^J1:[A-Z]+\d+$'
        $pivot.RowFields | Should -Contain 'Region'
        @($pivot.DataFields).Count | Should -BeGreaterThan 0
        $pivot.DataFields[0].FieldName | Should -Be 'Sales'
        $pivot.DataFields[0].DisplayName | Should -Be 'Total Sales'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets.Count | Should -Be 1
            $doc.Sheets[0].IsProtected | Should -BeTrue
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'uses a table-editing protection preset for protected worksheets' {
        $path = Join-Path $TestDrive 'DslExcelProtectedTableEditing.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Protect-OfficeExcelSheet -AllowTableEditing
            }
        }

        $protection = $null
        foreach ($entry in (Get-ZipEntriesLocal -Path $path | Where-Object { $_ -like 'xl/worksheets/*.xml' })) {
            $worksheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry $entry
            $protection = $worksheetXml.SelectSingleNode("//*[local-name()='sheetProtection']")
            if ($protection) {
                break
            }
        }

        $protection | Should -Not -BeNullOrEmpty
        $protection.GetAttribute('sheet') | Should -Match '^(1|true)$'
        $protection.GetAttribute('selectLockedCells') | Should -Match '^(0|false)$'
        $protection.GetAttribute('selectUnlockedCells') | Should -Match '^(0|false)$'
        $protection.GetAttribute('insertRows') | Should -Match '^(0|false)$'
        $protection.GetAttribute('sort') | Should -Match '^(0|false)$'
        $protection.GetAttribute('autoFilter') | Should -Match '^(0|false)$'
    }

    It 'lists and clears conditional formatting and data validation rules' {
        $path = Join-Path $TestDrive 'DslExcelRuleManagement.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100; Rate = 0.2 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200; Rate = 0.45 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150; Rate = 0.33 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelConditionalRule -Range 'B2:B4' -Operator GreaterThan -Formula1 '150'
                Add-OfficeExcelConditionalColorScale -Range 'C2:C4' -StartColor '#FEE599' -EndColor '#6AA84F'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B4' -Operator Between -Formula1 1 -Formula2 1000 -AllowBlank:$false
                Add-OfficeExcelValidationDecimal -Range 'C2:C4' -Operator Between -Formula1 0.0 -Formula2 1.0
            }
        }

        $conditionalRules = @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data)
        $conditionalRules.Count | Should -Be 2
        ($conditionalRules | Where-Object Range -EQ 'B2:B4').Operator | Should -Be 'GreaterThan'
        ($conditionalRules | Where-Object Range -EQ 'B2:B4').Formulas | Should -Contain '150'

        $rangeFilteredRules = @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data -Range 'B3')
        $rangeFilteredRules.Count | Should -Be 1
        $rangeFilteredRules[0].Range | Should -Be 'B2:B4'

        $validations = @(Get-OfficeExcelDataValidation -Path $path -Sheet Data)
        $validations.Count | Should -Be 2
        ($validations | Where-Object Range -EQ 'B2:B4').Formula1 | Should -Be '1'
        ($validations | Where-Object Range -EQ 'B2:B4').Formula2 | Should -Be '1000'

        Clear-OfficeExcelConditionalFormatting -Path $path -Sheet Data -Range 'B3' -Confirm:$false
        Clear-OfficeExcelDataValidation -Path $path -Sheet Data -Range 'C3' -Confirm:$false

        @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data -Range 'B3').Count | Should -Be 0
        @(Get-OfficeExcelDataValidation -Path $path -Sheet Data -Range 'C3').Count | Should -Be 0

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'sqref="B2 B4"'
        $worksheetXml | Should -Match 'sqref="C2 C4"'

        Clear-OfficeExcelConditionalFormatting -Path $path -Sheet Data -Confirm:$false
        Clear-OfficeExcelDataValidation -Path $path -Sheet Data -Confirm:$false

        @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data).Count | Should -Be 0
        @(Get-OfficeExcelDataValidation -Path $path -Sheet Data).Count | Should -Be 0
    }

    It 'reads conditional formatting rules from the current DSL sheet by default' {
        $path = Join-Path $TestDrive 'DslExcelConditionalFormattingContextRead.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 1
                Set-OfficeExcelCell -Address 'A2' -Value 2
                Add-OfficeExcelConditionalRule -Range 'A1:A2' -Operator GreaterThan -Formula1 '1'

                $rules = @(Get-OfficeExcelConditionalFormatting -Range 'A2')
                $rules.Count | Should -Be 1
                $rules[0].SheetName | Should -Be 'Data'
                $rules[0].Range | Should -Be 'A1:A2'
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'reads data validation rules from the current DSL sheet by default' {
        $path = Join-Path $TestDrive 'DslExcelDataValidationContextRead.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'B2' -Value 10
                Set-OfficeExcelCell -Address 'B3' -Value 20
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B4' -Operator Between -Formula1 1 -Formula2 100

                $validations = @(Get-OfficeExcelDataValidation -Range 'B3')
                $validations.Count | Should -Be 1
                $validations[0].SheetName | Should -Be 'Data'
                $validations[0].Range | Should -Be 'B2:B4'
            }
        }

        Test-Path $path | Should -BeTrue
    }

    It 'adds friendly conditional formatting rule types through the thin DSL surface' {
        $path = Join-Path $TestDrive 'DslExcelConditionalRuleTypes.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Ready'
                Set-OfficeExcelCell -Address 'A2' -Value 'Blocked'
                Set-OfficeExcelCell -Address 'A3' -Value 'Ready'
                Set-OfficeExcelCell -Address 'B1' -Value 1
                Set-OfficeExcelCell -Address 'B2' -Value 2
                Set-OfficeExcelCell -Address 'B3' -Value 3
                Set-OfficeExcelCell -Address 'C1' -Value ([datetime]'2026-06-22')
                Set-OfficeExcelCell -Address 'C2' -Value ([datetime]'2026-06-21')

                Add-OfficeExcelConditionalRule -Range 'A1:A3' -RuleType DuplicateValues
                Add-OfficeExcelConditionalRule -Range 'A1:A3' -RuleType UniqueValues
                Add-OfficeExcelConditionalRule -Range 'A1:A3' -RuleType ContainsText -Text Ready
                Add-OfficeExcelConditionalRule -Range 'A1:A3' -RuleType BeginsWith -Text Blo
                Add-OfficeExcelConditionalRule -Range 'B1:B3' -RuleType top10 -Rank 1
                Add-OfficeExcelConditionalRule -Range 'B1:B3' -RuleType BelowAverage -EqualAverage -StandardDeviation 1
                Add-OfficeExcelConditionalRule -Range 'D1:D3' -RuleType ContainsBlanks
                Add-OfficeExcelConditionalRule -Range 'E1:E3' -RuleType NotContainsErrors
                Add-OfficeExcelConditionalRule -Range 'C1:C3' -RuleType TimePeriod -TimePeriod Today
                Add-OfficeExcelConditionalRule -Range 'B1:B3' -RuleType formula -Formula1 'B1>1' -StopIfTrue
            }
        }

        $rules = @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data)
        $rules.Count | Should -Be 10
        $rules.Type | Should -Contain 'DuplicateValues'
        $rules.Type | Should -Contain 'UniqueValues'
        $rules.Type | Should -Contain 'ContainsText'
        $rules.Type | Should -Contain 'BeginsWith'
        $rules.Type | Should -Contain 'Top10'
        $rules.Type | Should -Contain 'AboveAverage'
        $rules.Type | Should -Contain 'ContainsBlanks'
        $rules.Type | Should -Contain 'NotContainsErrors'
        $rules.Type | Should -Contain 'TimePeriod'
        $rules.Type | Should -Contain 'Expression'
        ($rules | Where-Object Type -EQ 'ContainsText').Formulas | Should -Contain 'NOT(ISERROR(SEARCH("Ready",A1)))'
        ($rules | Where-Object Type -EQ 'Expression').StopIfTrue | Should -BeTrue
    }

    It 'adds Power Query metadata from the current Excel DSL sheet' {
        $path = Join-Path $TestDrive 'DslExcelPowerQueryContext.xlsx'
        $script:PowerQueryContextMetadata = $null

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Region'
                Set-OfficeExcelCell -Address 'A2' -Value 'EU'
                $script:PowerQueryContextMetadata = Add-OfficeExcelPowerQueryMetadata -Name 'ContextQuery' -CommandText 'let Source = Excel.CurrentWorkbook() in Source' -PassThru
            }
        }

        $script:PowerQueryContextMetadata.AddedWorkbookConnection | Should -BeTrue
        $script:PowerQueryContextMetadata.AddedWorksheetQueryTable | Should -BeTrue
        $script:PowerQueryContextMetadata.QueryTableName | Should -Be 'ContextQueryTable'

        $dataModel = Get-OfficeExcelDataModel -Path $path
        $dataModel.HasDataModelOrQueries | Should -BeTrue
    }

    It 'targets conditional formatting and validation by table header' {
        $path = Join-Path $TestDrive 'DslExcelHeaderTargets.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100; Status = 'New' }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200; Status = 'Done' }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150; Status = 'New' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelConditionalRule -HeaderName Sales -TableName Sales -Operator GreaterThan -Formula1 '150' -PassThru | Should -Be 'B2:B4'
                Add-OfficeExcelConditionalColorScale -ColumnName Sales -TableName Sales -StartColor '#FEE599' -EndColor '#6AA84F' -PassThru | Should -Be 'B2:B4'
                Add-OfficeExcelValidationList -HeaderName Status -TableName Sales -Values 'New','Done' -PassThru | Should -Be 'C2:C4'
            }
        }

        $conditionalRules = @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data -HeaderName Sales -TableName Sales)
        ($conditionalRules | Where-Object Range -EQ 'B2:B4').Count | Should -Be 2

        $validations = @(Get-OfficeExcelDataValidation -Path $path -Sheet Data -ColumnName Status -TableName Sales)
        @($validations | Where-Object Range -EQ 'C2:C4').Count | Should -Be 1

        $updated = @(Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -HeaderName Status -TableName Sales -PromptTitle Status -Prompt 'Pick a listed status' -ErrorTitle Invalid -ErrorMessage 'Use the list' -PassThru)
        $updated.Count | Should -Be 1
        $updated[0].Range | Should -Be 'C2:C4'
        $updated[0].PromptTitle | Should -Be 'Status'

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'sqref="B2:B4"'
        $worksheetXml | Should -Match 'sqref="C2:C4"'
        $worksheetXml | Should -Match 'promptTitle="Status"'

        Clear-OfficeExcelConditionalFormatting -Path $path -Sheet Data -HeaderName Sales -TableName Sales -Confirm:$false
        Clear-OfficeExcelDataValidation -Path $path -Sheet Data -ColumnName Status -TableName Sales -Confirm:$false

        @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data -HeaderName Sales -TableName Sales).Count | Should -Be 0
        @(Get-OfficeExcelDataValidation -Path $path -Sheet Data -HeaderName Status -TableName Sales).Count | Should -Be 0
    }

    It 'sets prompt and error messages on existing data validation rules' {
        $path = Join-Path $TestDrive 'DslExcelDataValidationMessages.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B4' -Operator Between -Formula1 1 -Formula2 1000
            }
        }

        $updated = @(Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B2:B4' -PromptTitle 'Sales input' -Prompt 'Use 1-1000' -ErrorTitle 'Invalid sales' -ErrorMessage 'Use 1-1000' -PassThru)
        $updated.Count | Should -Be 1
        $updated[0].PromptTitle | Should -Be 'Sales input'
        $updated[0].Prompt | Should -Be 'Use 1-1000'
        $updated[0].ErrorTitle | Should -Be 'Invalid sales'
        $updated[0].Error | Should -Be 'Use 1-1000'

        $validation = Get-OfficeExcelDataValidation -Path $path -Sheet Data -Range 'B3'
        $validation.PromptTitle | Should -Be 'Sales input'
        $validation.Prompt | Should -Be 'Use 1-1000'
        $validation.ErrorTitle | Should -Be 'Invalid sales'
        $validation.Error | Should -Be 'Use 1-1000'

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'promptTitle="Sales input"'
        $worksheetXml | Should -Match 'prompt="Use 1-1000"'
        $worksheetXml | Should -Match 'errorTitle="Invalid sales"'
        $worksheetXml | Should -Match 'error="Use 1-1000"'
        $worksheetXml | Should -Match 'showInputMessage="(?:1|true)"'
        $worksheetXml | Should -Match 'showErrorMessage="(?:1|true)"'

        $updated = @(Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B2:B4' -Prompt 'Use a whole number' -PassThru)
        $updated.Count | Should -Be 1
        $updated[0].PromptTitle | Should -Be 'Sales input'
        $updated[0].Prompt | Should -Be 'Use a whole number'
        $updated[0].ErrorTitle | Should -Be 'Invalid sales'
        $updated[0].Error | Should -Be 'Use 1-1000'

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'promptTitle="Sales input"'
        $worksheetXml | Should -Match 'prompt="Use a whole number"'
        $worksheetXml | Should -Match 'errorTitle="Invalid sales"'
        $worksheetXml | Should -Match 'error="Use 1-1000"'
        $worksheetXml | Should -Match 'showInputMessage="(?:1|true)"'
        $worksheetXml | Should -Match 'showErrorMessage="(?:1|true)"'

        $hiddenMessageXml = $worksheetXml -replace 'showInputMessage="(?:1|true)"', 'showInputMessage="0"' -replace 'showErrorMessage="(?:1|true)"', 'showErrorMessage="0"'
        Set-XlsxEntryTextLocal -Path $path -Entry 'xl/worksheets/sheet1.xml' -Text $hiddenMessageXml

        $updated = @(Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B2:B4' -Prompt 'Keep prompt text hidden' -PassThru)
        $updated.Count | Should -Be 1
        $updated[0].PromptTitle | Should -Be 'Sales input'
        $updated[0].Prompt | Should -Be 'Keep prompt text hidden'
        $updated[0].ErrorTitle | Should -Be 'Invalid sales'
        $updated[0].Error | Should -Be 'Use 1-1000'

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 'promptTitle="Sales input"'
        $worksheetXml | Should -Match 'prompt="Keep prompt text hidden"'
        $worksheetXml | Should -Match 'errorTitle="Invalid sales"'
        $worksheetXml | Should -Match 'error="Use 1-1000"'
        $worksheetXml | Should -Match 'showInputMessage="(?:0|false)"'
        $worksheetXml | Should -Match 'showErrorMessage="(?:0|false)"'
    }

    It 'preserves omitted prompt and error fields per matched data validation rule' {
        $path = Join-Path $TestDrive 'DslExcelDataValidationMessages.PerRule.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B3' -Operator Between -Formula1 1 -Formula2 1000
                Add-OfficeExcelValidationWholeNumber -Range 'B4:B4' -Operator Between -Formula1 1 -Formula2 1000
            }
        }

        Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B2:B3' -PromptTitle 'North sales' -Prompt 'Use north sales' -ErrorTitle 'North invalid' -ErrorMessage 'North error'
        Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B4:B4' -PromptTitle 'South sales' -Prompt 'Use south sales' -ErrorTitle 'South invalid' -ErrorMessage 'South error'

        $updated = @(Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B2:B4' -Prompt 'Shared prompt' -PassThru)
        $updated.Count | Should -Be 2

        $firstRule = @(Get-OfficeExcelDataValidation -Path $path -Sheet Data -Range 'B2' | Where-Object Range -EQ 'B2:B3')[0]
        $secondRule = @(Get-OfficeExcelDataValidation -Path $path -Sheet Data -Range 'B4' | Where-Object Range -EQ 'B4:B4')[0]

        $firstRule.PromptTitle | Should -Be 'North sales'
        $firstRule.Prompt | Should -Be 'Shared prompt'
        $firstRule.ErrorTitle | Should -Be 'North invalid'
        $firstRule.Error | Should -Be 'North error'

        $secondRule.PromptTitle | Should -Be 'South sales'
        $secondRule.Prompt | Should -Be 'Shared prompt'
        $secondRule.ErrorTitle | Should -Be 'South invalid'
        $secondRule.Error | Should -Be 'South error'
    }

    It 'preserves data validation display flags per matched rule' {
        $path = Join-Path $TestDrive 'DslExcelDataValidationMessages.PerRuleFlags.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B3' -Operator Between -Formula1 1 -Formula2 1000
                Add-OfficeExcelValidationWholeNumber -Range 'B4:B4' -Operator Between -Formula1 1 -Formula2 1000
            }
        }

        Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B2:B3' -PromptTitle 'Hidden sales' -Prompt 'Keep this hidden' -ShowInputMessage:$false

        $updated = @(Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B2:B4' -Prompt 'Shared prompt' -PassThru)
        $updated.Count | Should -Be 2

        $worksheet = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $validations = @($worksheet.GetElementsByTagName('dataValidation', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'))
        $hiddenRule = @($validations | Where-Object { $_.sqref -eq 'B2:B3' })[0]
        $newMessageRule = @($validations | Where-Object { $_.sqref -eq 'B4:B4' })[0]

        $hiddenRule.promptTitle | Should -Be 'Hidden sales'
        $hiddenRule.prompt | Should -Be 'Shared prompt'
        $hiddenRule.showInputMessage | Should -Match '^(0|false)$'
        $newMessageRule.promptTitle | Should -BeNullOrEmpty
        $newMessageRule.prompt | Should -Be 'Shared prompt'
        $newMessageRule.showInputMessage | Should -Match '^(1|true)$'
    }

    It 'updates message display state for full-column and full-row validation ranges' {
        $path = Join-Path $TestDrive 'DslExcelDataValidationMessages.FullRanges.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100; Status = 'New' }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200; Status = 'Done' }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150; Status = 'New' }
            [PSCustomObject]@{ Region = 'LATAM'; Sales = 250; Status = 'Done' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B5' -Operator Between -Formula1 1 -Formula2 1000
                Add-OfficeExcelValidationList -Range 'C5:D5' -Values 'New','Done'
            }
        }

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml = $worksheetXml -replace 'sqref="B2:B5"', 'sqref="B:B"'
        $worksheetXml = $worksheetXml -replace 'sqref="C5:D5"', 'sqref="5:5"'
        Set-XlsxEntryTextLocal -Path $path -Entry 'xl/worksheets/sheet1.xml' -Text $worksheetXml

        Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range 'B:B' -Prompt 'Whole column prompt'
        Set-OfficeExcelDataValidationMessage -Path $path -Sheet Data -Range '5:5' -ErrorMessage 'Whole row error'

        $worksheet = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $validations = @($worksheet.GetElementsByTagName('dataValidation', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'))
        $columnRule = @($validations | Where-Object { $_.sqref -eq 'B:B' })[0]
        $rowRule = @($validations | Where-Object { $_.sqref -eq '5:5' })[0]

        $columnRule.prompt | Should -Be 'Whole column prompt'
        $columnRule.showInputMessage | Should -Match '^(1|true)$'
        $rowRule.error | Should -Be 'Whole row error'
        $rowRule.showErrorMessage | Should -Match '^(1|true)$'
    }

    It 'clears range contents and attached metadata through a thin range command' {
        $path = Join-Path $TestDrive 'DslExcelRangeClear.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100; Rate = 0.2 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200; Rate = 0.45 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150; Rate = 0.33 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelConditionalRule -Range 'B2:B4' -Operator GreaterThan -Formula1 '150'
                Add-OfficeExcelValidationWholeNumber -Range 'B2:B4' -Operator Between -Formula1 1 -Formula2 1000
            }
        }

        Clear-OfficeExcelRange -Path $path -Sheet Data -Range 'B2:B4' -Contents -DataValidations -ConditionalFormatting -Confirm:$false

        $clearedRows = @(Get-OfficeExcelRange -Path $path -Sheet Data -Range 'A1:C4')
        $clearedRows.Count | Should -Be 3
        foreach ($row in $clearedRows) {
            $row.Sales | Should -BeNullOrEmpty
            $row.Rate | Should -Not -BeNullOrEmpty
        }

        @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data -Range 'B3').Count | Should -Be 0
        @(Get-OfficeExcelDataValidation -Path $path -Sheet Data -Range 'B3').Count | Should -Be 0
    }

    It 'sets and reads mixed rich text runs in Excel cells' {
        Get-Command ExcelNew | Should -Not -BeNullOrEmpty
        Get-Command ExcelTextRun | Should -Not -BeNullOrEmpty

        $path = Join-Path $TestDrive 'DslExcelRichText.xlsx'

        ExcelNew -Path $path {
            ExcelSheet -Name 'Summary' -Content {
                ExcelRichText -Address A1 -Run @(
                    ExcelTextRun 'Status: '
                    ExcelTextRun 'Blocked' -Bold -Color Red -FontName 'Arial' -FontSize 14
                    ExcelTextRun ' pending owner' -Italic -Underline
                )
            }
        }

        $runs = @(Get-OfficeExcelRichText -Path $path -Sheet Summary -Address A1)
        $runs.Count | Should -Be 3
        $runs[0].Text | Should -Be 'Status: '
        $runs[1].Text | Should -Be 'Blocked'
        $runs[1].Bold | Should -BeTrue
        $runs[1].FontColor | Should -Match 'FF0000'
        $runs[1].FontName | Should -Be 'Arial'
        $runs[1].FontSize | Should -Be 14
        $runs[2].Italic | Should -BeTrue
        $runs[2].Underline | Should -BeTrue

        $roundTripPath = Join-Path $TestDrive 'DslExcelRichText.RoundTrip.xlsx'
        ExcelNew -Path $roundTripPath {
            ExcelSheet -Name 'Summary' -Content {
                ExcelRichText -Address A1 -Run $runs
            }
        }

        $roundTripXml = Read-XlsxEntryText -Path $roundTripPath -Entry 'xl/worksheets/sheet1.xml'
        $roundTripXml | Should -Match 'rgb="FFFF0000"'
        $roundTripXml | Should -Not -Match 'rgb="FFFFFF00"'

        $contextRuns = @()
        New-OfficeExcel -Path (Join-Path $TestDrive 'DslExcelRichText.Context.xlsx') {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Set-OfficeExcelRichText -Address A1 -Run @(
                    'Context '
                    @{ Text = 'read'; Bold = $true }
                )
                $contextRuns = @(Get-OfficeExcelRichText -Address A1)
                $contextRuns.Count | Should -Be 2
                $contextRuns[0].Text | Should -Be 'Context '
                $contextRuns[1].Bold | Should -BeTrue
            }
        }

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match 't="inlineStr"'
        $worksheetXml | Should -Match '<(?:\w+:)?rPr>'
        [xml] $worksheetDocument = $worksheetXml
        @($worksheetDocument.SelectNodes("//*[local-name()='rPr']/*[local-name()='b' and (not(@val) or @val='1' or @val='true')]")).Count | Should -BeGreaterThan 0
        @($worksheetDocument.SelectNodes("//*[local-name()='rPr']/*[local-name()='i' and (not(@val) or @val='1' or @val='true')]")).Count | Should -BeGreaterThan 0
        @($worksheetDocument.SelectNodes("//*[local-name()='rPr']/*[local-name()='u' and (not(@val) or @val='single')]")).Count | Should -BeGreaterThan 0
        $worksheetXml | Should -Match 'rgb="FFFF0000"'
    }

    It 'protects and unprotects workbook structure through thin workbook commands' {
        $path = Join-Path $TestDrive 'DslExcelWorkbookProtection.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
            }
            Protect-OfficeExcelWorkbook -LegacyPasswordHash 'CAFE' -ProtectWindows
        }

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.IsWorkbookProtected | Should -BeTrue
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $workbookXml = Read-XlsxEntryText -Path $path -Entry 'xl/workbook.xml'
        $workbookXml | Should -Match '<(?:\w+:)?workbookProtection\b'
        $workbookXml | Should -Match 'lockStructure="1"'
        $workbookXml | Should -Match 'lockWindows="1"'
        $workbookXml | Should -Match 'workbookPassword="CAFE"'

        Unprotect-OfficeExcelWorkbook -Path $path

        $protectedFile = Protect-OfficeExcelWorkbook -Path $path -LegacyPasswordHash 'CAFE' -PassThru
        $protectedFile | Should -BeOfType ([System.IO.FileInfo])
        $protectedFile.FullName | Should -Be (Resolve-Path -LiteralPath $path).Path
        $unprotectedFile = Unprotect-OfficeExcelWorkbook -Path $path -PassThru
        $unprotectedFile | Should -BeOfType ([System.IO.FileInfo])
        $unprotectedFile.FullName | Should -Be (Resolve-Path -LiteralPath $path).Path

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.IsWorkbookProtected | Should -BeFalse
        } finally {
            Close-OfficeExcel -Document $doc
        }

        (Read-XlsxEntryText -Path $path -Entry 'xl/workbook.xml') | Should -Not -Match '<workbookProtection\b'
    }

    It 'renames and removes named ranges without saving unless requested' {
        $path = Join-Path $TestDrive 'DslExcelNamedRangeSaveSemantics.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Name', 'Value'
                Set-OfficeExcelNamedRange -Name 'Totals' -Range 'A1:B1'
            }
        }

        $doc = Get-OfficeExcel -Path $path
        try {
            ($doc | Rename-OfficeExcelNamedRange -Name 'Totals' -NewName 'GrandTotal' -Sheet 'Data' -PassThru) | Should -BeTrue
        } finally {
            Close-OfficeExcel -Document $doc
        }

        (Get-OfficeExcelNamedRange -Path $path -Sheet 'Data').Name | Should -Contain 'Totals'
        (Get-OfficeExcelNamedRange -Path $path -Sheet 'Data').Name | Should -Not -Contain 'GrandTotal'

        $doc = Get-OfficeExcel -Path $path
        try {
            ($doc | Rename-OfficeExcelNamedRange -Name 'Totals' -NewName 'GrandTotal' -Sheet 'Data' -Save -PassThru) | Should -BeTrue
            ($doc | Remove-OfficeExcelNamedRange -Name 'GrandTotal' -Sheet 'Data' -PassThru) | Should -BeTrue
        } finally {
            Close-OfficeExcel -Document $doc
        }

        (Get-OfficeExcelNamedRange -Path $path -Sheet 'Data').Name | Should -Contain 'GrandTotal'

        $doc = Get-OfficeExcel -Path $path
        try {
            ($doc | Remove-OfficeExcelNamedRange -Name 'GrandTotal' -Sheet 'Data' -Save -PassThru) | Should -BeTrue
        } finally {
            Close-OfficeExcel -Document $doc
        }

        @(Get-OfficeExcelNamedRange -Path $path -Sheet 'Data' | Where-Object Name -eq 'GrandTotal').Count | Should -Be 0
    }

    It 'adds, lists, and clears manual worksheet page breaks' {
        $path = Join-Path $TestDrive 'DslExcelPageBreaks.xlsx'
        $rows = 1..12 | ForEach-Object {
            [PSCustomObject]@{
                Name = "Item $_"
                Value = $_
            }
        }

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Items'
                Add-OfficeExcelPageBreak -Row 5 -Column 1
                $script:contextPageBreaks = @(Get-OfficeExcelPageBreak)
            }
        }

        $breaks = @(Get-OfficeExcelPageBreak -Path $path -Sheet Data)
        $breaks.Count | Should -Be 2
        ($breaks | Where-Object Type -EQ 'Row').Index | Should -Be 5
        ($breaks | Where-Object Type -EQ 'Row').Position | Should -Be 5
        ($breaks | Where-Object Type -EQ 'Column').Index | Should -Be 1
        ($breaks | Where-Object Type -EQ 'Column').Position | Should -Be 1

        $bothFilters = @(Get-OfficeExcelPageBreak -Path $path -Sheet Data -Row -Column)
        $bothFilters.Count | Should -Be 2

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml | Should -Match '<(?:\w+:)?rowBreaks\b'
        $worksheetXml | Should -Match '<(?:\w+:)?colBreaks\b'
        $worksheetXml | Should -Match 'id="5"'
        $worksheetXml | Should -Match 'id="1"'

        $beforeWhatIf = [System.IO.File]::ReadAllBytes($path)
        Clear-OfficeExcelPageBreak -Path $path -Sheet Data -Row 5 -WhatIf
        $afterWhatIf = [System.IO.File]::ReadAllBytes($path)
        [Convert]::ToBase64String($afterWhatIf) | Should -Be ([Convert]::ToBase64String($beforeWhatIf))

        Clear-OfficeExcelPageBreak -Path $path -Sheet Data -Row 5 -Confirm:$false

        $rowBreaks = @(Get-OfficeExcelPageBreak -Path $path -Sheet Data -Row)
        $rowBreaks.Count | Should -Be 0
        $columnBreaks = @(Get-OfficeExcelPageBreak -Path $path -Sheet Data -Column)
        $columnBreaks.Count | Should -Be 1
        $columnBreaks[0].Index | Should -Be 1

        $script:contextPageBreaks.Count | Should -Be 2
        $script:contextPageBreaks[0].SheetName | Should -Be 'Data'

        Clear-OfficeExcelPageBreak -Path $path -Sheet Data -All -Confirm:$false

        @(Get-OfficeExcelPageBreak -Path $path -Sheet Data).Count | Should -Be 0
    }

    It 'uses OfficeIMO pivot field options and captions' {
        $path = Join-Path $TestDrive 'DslExcelPivotOptions.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Product = 'Standard'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Product = 'Standard'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Product = 'Legacy'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelPivotTable -SourceRange 'A1:C4' -DestinationCell 'E1' -RowField 'Region' -PageField 'Product' -DataField 'Sales' -DataNumberFormat '#,##0' -GrandTotalCaption 'Overall' -FieldSort @{ Region = 'Ascending' } -FieldHiddenItems @{ Region = @('APAC') } -PageFieldSelection @{ Product = 'Standard' } -RefreshOnOpen -NoSaveSourceData -NoPreserveFormatting -DisableDrill
            }
        }

        Test-Path $path | Should -BeTrue

        $pivot = @(Get-OfficeExcelPivotTable -Path $path)[0]
        $pivot.RefreshOnOpen | Should -BeTrue
        $pivot.SaveSourceData | Should -BeFalse
        $pivot.PreserveFormatting | Should -BeFalse
        $pivot.EnableDrill | Should -BeFalse
    }

    It 'configures refresh on open across pivot caches and connection metadata' {
        $path = Join-Path $TestDrive 'DslExcelRefreshOnOpen.xlsx'
        $connectionXml = '<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><connection id="1" name="SalesConnection" type="5" refreshedVersion="7"/></connections>'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelPivotTable -SourceRange 'A1:B4' -DestinationCell 'E1' -Name 'SalesPivot' -RowField 'Region' -DataField 'Sales' -NoRefreshOnOpen -SaveSourceData
                Add-OfficeExcelPackageMetadata -Kind Connection -Xml $connectionXml
                $result = Set-OfficeExcelRefreshOnOpen -NoSavePivotSourceData -PassThru
                $result.Enabled | Should -BeTrue
                $result.PivotCacheCount | Should -Be 1
                $result.ConnectionCount | Should -Be 1
            }
        }

        $pivot = @(Get-OfficeExcelPivotTable -Path $path)[0]
        $pivot.RefreshOnOpen | Should -BeTrue
        $pivot.SaveSourceData | Should -BeFalse

        $connectionText = $null
        foreach ($entry in (Get-ZipEntriesLocal -Path $path | Where-Object { $_ -notlike '*.rels' -and $_ -ne '[Content_Types].xml' })) {
            $text = Read-XlsxEntryText -Path $path -Entry $entry
            if ($text -match '<(?:\w+:)?connections\b') {
                $connectionText = $text
                break
            }
        }

        $connectionText | Should -Not -BeNullOrEmpty
        $connectionText | Should -Match 'refreshOnLoad="1"'
    }

    It 'targets pivot output ranges for conditional formatting' {
        $path = Join-Path $TestDrive 'DslExcelPivotConditionalFormatting.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelPivotTable -SourceRange 'A1:B4' -DestinationCell 'E1' -Name 'SalesPivot' -RowField 'Region' -DataField 'Sales'
                Add-OfficeExcelConditionalRule -PivotTableName 'SalesPivot' -Operator GreaterThan -Formula1 '0' -PassThru | Should -Be 'F2:F2'
                Add-OfficeExcelConditionalColorScale -PivotTableName 'SalesPivot' -StartColor '#FEE599' -EndColor '#6AA84F' -PassThru | Should -Be 'F2:F2'
                Add-OfficeExcelConditionalDataBar -PivotTableName 'SalesPivot' -Color '#92D050' -PassThru | Should -Be 'F2:F2'
                Add-OfficeExcelConditionalIconSet -PivotTableName 'SalesPivot' -PassThru | Should -Be 'F2:F2'
            }
        }

        $conditionalRules = @(Get-OfficeExcelConditionalFormatting -Path $path -Sheet Data -Range 'F2')
        $conditionalRules.Count | Should -Be 4

        $worksheetXml = Read-XlsxEntryText -Path $path -Entry 'xl/worksheets/sheet1.xml'
        ($worksheetXml | Select-String 'sqref="F2:F2"' -AllMatches).Matches.Count | Should -Be 4
    }

    It 'saves pivot and sparkline workbooks with reopenable package parts' {
        $path = Join-Path $TestDrive 'DslExcelPivotSparklineOpen.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
                Add-OfficeExcelPivotTable -SourceRange 'A1:B4' -DestinationCell 'E1' -RowField 'Region' -DataField 'Sales'
                Add-OfficeExcelSparkline -DataRange 'B2:B4' -LocationRange 'D2:D4' -Type Line
            }
        }

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $summary = Get-OfficeExcelSummary -Document $doc -IncludeSchema
            $summary.PivotTableCount | Should -Be 1
            $summary.SparklineGroupCount | Should -Be 1
            $summary.Schema.Worksheets[0].TableCount | Should -Be 1
            $summary.Schema.Tables[0].Columns | Should -Contain 'Region'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $contentTypes = Get-ZipXmlDocumentLocal -Path $path -Entry '[Content_Types].xml'
        $contentTypes.OuterXml | Should -Match 'pivotTable'
        $workbookXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/workbook.xml'
        $workbookXml.OuterXml | Should -Match 'Data'
        $worksheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $worksheetXml.OuterXml | Should -Match 'sparklineGroup'

        $pivotTables = @(Get-OfficeExcelPivotTable -Path $path)
        $pivotTables.Count | Should -Be 1
        $pivotTables[0].RowFields | Should -Contain 'Region'
    }

    It 'supports advanced Excel page setup and visibility helpers' {
        $path = Join-Path $TestDrive 'DslExcelAdvancedLayout.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                Set-OfficeExcelPageSetup -FitToWidth 1 -FitToHeight 0
                Set-OfficeExcelMargins -Preset Narrow
                Set-OfficeExcelOrientation -Orientation Landscape
                Set-OfficeExcelGridlines -Hide
                Set-OfficeExcelFreeze -TopRows 1 -LeftColumns 1
                Set-OfficeExcelSheetVisibility -Hide
            }
        }

        Test-Path $path | Should -BeTrue

        $workbookXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/workbook.xml'
        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'

        $workbookSheet = $workbookXml.SelectSingleNode("/*[local-name()='workbook']/*[local-name()='sheets']/*[local-name()='sheet']")
        $workbookSheet.GetAttribute('name') | Should -Be 'Data'
        $workbookSheet.GetAttribute('state') | Should -Be 'hidden'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.HiddenSheetCount | Should -Be 1
        $summary.Sheets[0].State | Should -Be 'Hidden'

        $view = Get-OfficeExcelWorksheetView -Path $path -Sheet Data
        $view.ShowGridlines | Should -BeFalse
        $view.FrozenRowCount | Should -Be 1
        $view.FrozenColumnCount | Should -Be 1
        $view.TopLeftCell | Should -Be 'B2'

        $pageSetup = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='pageSetup']")
        $pageSetup.GetAttribute('fitToWidth') | Should -Be '1'
        $pageSetup.GetAttribute('fitToHeight') | Should -Be '0'
        $pageSetup.GetAttribute('orientation') | Should -Be 'landscape'

        $pageMargins = $sheetXml.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='pageMargins']")
        $pageMargins.GetAttribute('left') | Should -Be '0.25'
        $pageMargins.GetAttribute('right') | Should -Be '0.25'
        $pageMargins.GetAttribute('top') | Should -Be '0.5'
        $pageMargins.GetAttribute('bottom') | Should -Be '0.5'
    }

    It 'wraps OfficeIMO worksheet operations and print definitions' {
        $path = Join-Path $TestDrive 'ExcelWorksheetOperations.xlsx'
        $sourcePath = Join-Path $TestDrive 'ExcelWorksheetOperationsSource.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Sales = 200 }
        )
        $moreRows = @(
            [PSCustomObject]@{ Region = 'APAC'; Sales = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
            Add-OfficeExcelSheet -Name 'More' -Content {
                Add-OfficeExcelTable -InputObject $moreRows -TableName 'MoreSales'
            }
        }
        New-OfficeExcel -Path $sourcePath {
            Add-OfficeExcelSheet -Name 'External' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Name'
                Set-OfficeExcelCell -Address 'A2' -Value 'Imported'
            }
        }

        Copy-OfficeExcelSheet -Path $path -SourceSheet 'Data' -NewName 'DataCopy' -PassThru | Should -Not -BeNullOrEmpty
        Move-OfficeExcelSheet -Path $path -Sheet 'DataCopy' -Index 0
        Copy-OfficeExcelSheet -Path $path -SourcePath $sourcePath -SourceSheet 'External' -NewName 'ExternalCopy' -CopyMode Package -PassThru | Should -Not -BeNullOrEmpty
        $join = Join-OfficeExcelSheet -Path $path -TargetSheet 'Data' -SourceSheet 'More' -MatchColumnsByHeader
        Set-OfficeExcelPrintArea -Path $path -Sheet 'Data' -Range 'A1:B4'
        Set-OfficeExcelPrintTitles -Path $path -Sheet 'Data' -FirstRow 1 -LastRow 1

        $join.RowsCopied | Should -Be 1
        $join.TargetSheetName | Should -Be 'Data'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.Sheets[0].Name | Should -Be 'DataCopy'
        $summary.Sheets.Name | Should -Contain 'ExternalCopy'

        $external = @(Import-OfficeExcel -Path $path -WorksheetName 'ExternalCopy' -Range 'A1:A2')
        $external.Count | Should -Be 1
        $external[0].Name | Should -Be 'Imported'

        $merged = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:B4')
        $merged.Count | Should -Be 3
        $merged[2].Region | Should -Be 'APAC'

        $differences = @(Compare-OfficeExcelRange -Path $path -LeftSheet 'Data' -RightSheet 'DataCopy')
        $differences.Count | Should -BeGreaterThan 0

        $names = @(Get-OfficeExcelNamedRange -Path $path -Sheet 'Data')
        @($names | Where-Object Name -eq '_xlnm.Print_Area').Count | Should -Be 1
        @($names | Where-Object Name -eq '_xlnm.Print_Titles').Count | Should -Be 1
    }

    It 'merges workbooks through the package copy fast path' {
        $targetPath = Join-Path $TestDrive 'ExcelWorkbookPackageMerge.xlsx'
        $sourceAPath = Join-Path $TestDrive 'ExcelWorkbookPackageMerge-A.xlsx'
        $sourceBPath = Join-Path $TestDrive 'ExcelWorkbookPackageMerge-B.xlsx'

        New-OfficeExcel -Path $sourceAPath {
            Add-OfficeExcelSheet -Name 'First' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Name'
                Set-OfficeExcelCell -Address 'A2' -Value 'Alpha'
            }
        }
        New-OfficeExcel -Path $sourceBPath {
            Add-OfficeExcelSheet -Name 'Second' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Name'
                Set-OfficeExcelCell -Address 'A2' -Value 'Beta'
            }
        }

        $results = @(Join-OfficeExcelWorkbook -Path $targetPath -SourcePath @($sourceAPath, $sourceBPath) -CopyMode Package -SheetNamePrefix 'Merged')
        $results.Count | Should -Be 2
        $results[0].SheetCount | Should -Be 1
        $results[1].SheetCount | Should -Be 1
        $results[0].TargetSheets | Should -Contain 'MergedFirst'
        $results[1].TargetSheets | Should -Contain 'MergedSecond'

        $summary = Get-OfficeExcelSummary -Path $targetPath -IncludeSheets
        $summary.Sheets.Name | Should -Contain 'MergedFirst'
        $summary.Sheets.Name | Should -Contain 'MergedSecond'

        $first = @(Import-OfficeExcel -Path $targetPath -WorksheetName 'MergedFirst' -Range 'A1:A2')
        $second = @(Import-OfficeExcel -Path $targetPath -WorksheetName 'MergedSecond' -Range 'A1:A2')
        $first[0].Name | Should -Be 'Alpha'
        $second[0].Name | Should -Be 'Beta'
    }

    It 'finds, replaces, and edits Excel row values' {
        $path = Join-Path $TestDrive 'ExcelFindReplaceEditRows.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Ada'; Status = 'Draft'; Score = 12 }
            [PSCustomObject]@{ Name = 'Grace'; Status = 'Draft'; Score = 21 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'People'
            }
        }

        @(Find-OfficeExcel -Path $path -Sheet 'Data' -Text 'Draft').Count | Should -Be 2
        $numericMatch = @(Find-OfficeExcel -Path $path -Sheet 'Data' -Text '12' -Exact)
        $numericMatch.Count | Should -Be 1
        $numericMatch[0].Value | Should -BeOfType ([double])
        $numericMatch[0].Value | Should -Be 12
        Update-OfficeExcelText -Path $path -Sheet 'Data' -OldValue 'Draft' -NewValue 'Ready' -PassThru | Should -Be 2
        Update-OfficeExcelText -Path $path -Sheet 'Data' -OldValue '12' -NewValue 'Twelve' -PassThru | Should -Be 0
        Edit-OfficeExcelRow -Path $path -Sheet 'Data' -ScriptBlock {
            param($row)
            if ($row.CellByHeader('Name').Value -eq 'Ada') {
                $row.Set('Status', 'Done')
            }
        }

        $updated = @(Import-OfficeExcel -Path $path -WorksheetName 'Data' -Range 'A1:C3')
        $updated[0].Status | Should -Be 'Done'
        $updated[1].Status | Should -Be 'Ready'
        $updated[0].Score | Should -Be 12
        @(Find-OfficeExcel -Path $path -Sheet 'Data' -Text '^Done$' -Regex).Count | Should -Be 1
    }

    It 'updates text in sparse worksheets without scanning the rectangular used range' {
        $path = Join-Path $TestDrive 'ExcelSparseTextReplacement.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Draft'
                Set-OfficeExcelCell -Address 'XFD1048576' -Value 'Draft'
            }
        }

        $matches = @(Find-OfficeExcel -Path $path -Sheet 'Data' -Text 'Draft' -Exact)
        $matches.Count | Should -Be 2
        @($matches | ForEach-Object { $_.Address }) | Should -Contain 'A1'
        @($matches | ForEach-Object { $_.Address }) | Should -Contain 'XFD1048576'
        Update-OfficeExcelText -Path $path -Sheet 'Data' -OldValue 'Draft' -NewValue 'Ready' -PassThru | Should -Be 2
    }

    It 'counts threaded comments in workbook summaries' {
        $path = Join-Path $TestDrive 'DslExcelThreadedComments.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Sales = 100 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales'
            }
        }

        $spreadsheetDocumentType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Packaging.SpreadsheetDocument'
        $worksheetPartType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Packaging.WorksheetPart'
        $threadedPartType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Packaging.WorksheetThreadedCommentsPart'
        $threadedCommentsType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedComments'
        $threadedCommentType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedComment'
        $threadedCommentTextType = Get-TestLoadedType -Name 'DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments.ThreadedCommentText'

        $openMethod = $spreadsheetDocumentType.GetMethod('Open', [type[]] @([string], [bool]))
        $openArguments = [object[]] @($path.ToString(), $true)
        $document = $openMethod.Invoke($null, $openArguments)
        try {
            $worksheetPart = @($document.WorkbookPart.WorksheetParts)[0]
            $addPartMethod = $worksheetPartType.GetMethods() |
                Where-Object { $_.Name -eq 'AddNewPart' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 0 } |
                Select-Object -First 1
            $threadedPart = $addPartMethod.MakeGenericMethod($threadedPartType).Invoke($worksheetPart, @())
            $threadedComments = [Activator]::CreateInstance($threadedCommentsType)
            $threadedComment = [Activator]::CreateInstance($threadedCommentType)
            $threadedComment.Ref = 'A2'
            $threadedComment.PersonId = '{00000000-0000-0000-0000-000000000001}'
            $threadedComment.Id = '{00000000-0000-0000-0000-000000000002}'
            $threadedCommentTextConstructor = $threadedCommentTextType.GetConstructor([type[]] @([string]))
            $threadedComment.AppendChild($threadedCommentTextConstructor.Invoke([object[]] @('Modern note'))) | Out-Null
            $threadedComments.AppendChild($threadedComment) | Out-Null
            $threadedComments.Save($threadedPart)
        } finally {
            $document.Dispose()
        }

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.CommentCount | Should -Be 1
        $summary.Sheets[0].CommentCount | Should -Be 1
    }

    It 'adds a table of contents and reads ranges with the new Excel readers' {
        $path = Join-Path $TestDrive 'DslExcelNavigation.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                Set-OfficeExcelNamedRange -Name 'SalesData' -Range 'A1:B3'
            }
            Add-OfficeExcelSheet -Name 'Notes' -Content {
                Set-OfficeExcelRow -Row 1 -Values 'Label', 'Value'
                Set-OfficeExcelRow -Row 2 -Values 'Generated', 'Yes'
            }
        } | Out-Null

        $usedRange = Get-OfficeExcelUsedRange -Path $path -Sheet 'Data' -AsDataTable
        $usedRange.Rows.Count | Should -Be 2
        $usedRange.Columns[0].ColumnName | Should -Be 'Region'
        $usedRange.Rows[0]['Region'] | Should -Be 'NA'

        Add-OfficeExcelTableOfContents -Path $path -IncludeNamedRanges -AddBackLinks

        $tocRows = @(Get-OfficeExcelRange -Path $path -Sheet 'TOC' -Range 'A3:C5' -AsHashtable)
        $tocRows.Count | Should -Be 2
        $tocRows[0]['Sheet'] | Should -Be 'Data'
        $tocRows[0]['Named Ranges'] | Should -Match 'SalesData'
        $tocRows[1]['Sheet'] | Should -Be 'Notes'

        $noteRows = @(Get-OfficeExcelRange -Path $path -Sheet 'Notes' -Range 'A1:B2')
        $noteRows.Count | Should -Be 1
        $noteRows[0].Label | Should -Be 'Generated'
        $noteRows[0].Value | Should -Be 'Yes'

        $dataRows = @(Get-OfficeExcelRange -Path $path -Sheet 'Data' -Range 'A1:B3')
        $dataRows.Count | Should -Be 2
        $dataRows[0].Region | Should -Be 'NA'

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.SheetCount | Should -Be 3
        $summary.VisibleSheetCount | Should -Be 3
        $summary.TableCount | Should -Be 2
        $summary.NamedRangeCount | Should -Be 1
        $summary.HyperlinkCount | Should -BeGreaterThan 0
        $summary.Sheets.Name | Should -Contain 'Data'
        ($summary.Sheets | Where-Object Name -eq 'Data').UsedRange | Should -Be 'A1:B5'
        ($summary.Sheets | Where-Object Name -eq 'Data').Tables.Name | Should -Contain 'Sales'

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $doc.Sheets[0].Name | Should -Be 'TOC'

            $backLink = $null
            $doc['Data'].TryGetCellText(5, 1, [ref] $backLink) | Should -BeTrue
            $backLink | Should -Be "$([char]0x2190) TOC"
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'includes chartsheet charts in workbook summaries' {
        $path = Join-Path $TestDrive 'WorkbookWithChartSheet.xlsx'
        $archive = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            function Add-ZipTextEntry {
                param(
                    [Parameter(Mandatory)]
                    [System.IO.Compression.ZipArchive] $Archive,

                    [Parameter(Mandatory)]
                    [string] $EntryName,

                    [Parameter(Mandatory)]
                    [string] $Content
                )

                $entry = $Archive.CreateEntry($EntryName)
                $stream = $entry.Open()
                try {
                    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
                    try {
                        $writer.Write($Content)
                    } finally {
                        $writer.Dispose()
                    }
                } finally {
                    $stream.Dispose()
                }
            }

            Add-ZipTextEntry -Archive $archive -EntryName '[Content_Types].xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>
'@
            Add-ZipTextEntry -Archive $archive -EntryName '_rels/.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/workbook.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Revenue Chart" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/worksheets/sheet1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="str"><v>Region</v></c><c r="B1" t="str"><v>Revenue</v></c></row>
    <row r="2"><c r="A2" t="str"><v>EMEA</v></c><c r="B2"><v>42</v></c></row>
  </sheetData>
</worksheet>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/chartsheets/sheet1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <drawing r:id="rId1"/>
</chartsheet>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/chartsheets/_rels/sheet1.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/drawings/drawing1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="0" y="0"/><xdr:ext cx="6000000" cy="4000000"/>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/drawings/_rels/drawing1.xml.rels' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>
'@
            Add-ZipTextEntry -Archive $archive -EntryName 'xl/charts/chart1.xml' -Content @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea><c:layout/></c:plotArea></c:chart>
</c:chartSpace>
'@
        } finally {
            $archive.Dispose()
        }

        $summary = Get-OfficeExcelSummary -Path $path -IncludeSheets
        $summary.SheetCount | Should -Be 2
        $summary.ChartCount | Should -Be 1
        ($summary.Sheets | Where-Object Name -eq 'Revenue Chart').ChartCount | Should -Be 1
    }

    It 'formats Excel charts with legend, labels, and style presets' {
        $path = Join-Path $TestDrive 'DslExcelChartFormatting.xlsx'
        $rows = @(
            [PSCustomObject]@{ Region = 'NA'; Revenue = 100 }
            [PSCustomObject]@{ Region = 'EMEA'; Revenue = 200 }
            [PSCustomObject]@{ Region = 'APAC'; Revenue = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                $chart = Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type Pie -Title 'Revenue Mix' -PassThru
                $formattedChart = $chart |
                    Set-OfficeExcelChartLegend -Position Right -PassThru |
                    Set-OfficeExcelChartDataLabels -ShowValue $true -ShowPercent $true -Position OutsideEnd -NumberFormat '0.0%' -SourceLinked:$false -PassThru |
                    Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10 -PassThru

                $formattedChart | Should -Not -BeNullOrEmpty
            }
        } | Out-Null

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'xl/drawings/charts/style*.xml' }).Count | Should -BeGreaterThan 0
        ($entries | Where-Object { $_ -like 'xl/drawings/charts/colors*.xml' }).Count | Should -BeGreaterThan 0

        $chartXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/drawings/charts/chart1.xml'
        $legendPosition = $chartXml.SelectSingleNode("/*[local-name()='chartSpace']/*[local-name()='chart']/*[local-name()='legend']/*[local-name()='legendPos']")
        $legendPosition | Should -Not -BeNullOrEmpty
        $legendPosition.GetAttribute('val') | Should -Be 'r'

        $dataLabels = $chartXml.SelectSingleNode("//*[local-name()='dLbls']")
        $dataLabels | Should -Not -BeNullOrEmpty
        $dataLabels.SelectSingleNode("*[local-name()='showVal']").GetAttribute('val') | Should -Be '1'
        $dataLabels.SelectSingleNode("*[local-name()='showPercent']").GetAttribute('val') | Should -Be '1'
        $dataLabels.SelectSingleNode("*[local-name()='dLblPos']").GetAttribute('val') | Should -Be 'outEnd'

        $numberFormat = $dataLabels.SelectSingleNode("*[local-name()='numFmt']")
        $numberFormat | Should -Not -BeNullOrEmpty
        $numberFormat.GetAttribute('formatCode') | Should -Be '0.0%'
    }

    It 'formats Excel chart axes series and trendlines' {
        $path = Join-Path $TestDrive 'DslExcelChartAxisSeriesTrendline.xlsx'
        $rows = @(
            [PSCustomObject]@{ Month = 'Jan'; Revenue = 100 }
            [PSCustomObject]@{ Month = 'Feb'; Revenue = 200 }
            [PSCustomObject]@{ Month = 'Mar'; Revenue = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                $chart = Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type Line -Title 'Revenue Trend' -PassThru
                { $chart | Set-OfficeExcelChartSeries -SeriesIndex 0 -LineWidthPoints 1.5 -ErrorAction Stop } |
                    Should -Throw '*LineColor is required*'
                { $chart | Set-OfficeExcelChartPoint -SeriesIndex 0 -PointIndex 1 -LineWidthPoints 1.5 -ErrorAction Stop } |
                    Should -Throw '*LineColor is required*'
                $formattedChart = $chart |
                    Set-OfficeExcelChartAxis -CategoryTitle 'Month' -ValueTitle 'Revenue' -ValueNumberFormat '$#,##0' -SourceLinked:$false -ValueMinimum 0 -ValueMajorUnit 100 -ShowValueMinorGridlines -ValueGridlineColor '#D9EAD3' -GridlineWidthPoints 0.75 -PassThru |
                    Set-OfficeExcelChartSeries -SeriesIndex 0 -LineColor '#1F4E79' -LineWidthPoints 1.5 -MarkerStyle Circle -MarkerSize 6 -MarkerFillColor '#4472C4' -PassThru |
                    Set-OfficeExcelChartPoint -SeriesName 'Revenue' -PointIndex 1 -FillColor '#70AD47' -LineColor '#7030A0' -LineWidthPoints 1.25 -PassThru |
                    Set-OfficeExcelChartTrendline -SeriesIndex 0 -Type Linear -DisplayEquation -DisplayRSquared -LineColor '#C00000' -LineWidthPoints 1.25 -PassThru

                $formattedChart | Should -Not -BeNullOrEmpty
            }
        } | Out-Null

        $chartXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/drawings/charts/chart1.xml'
        $chartOuterXml = $chartXml.OuterXml

        $categoryAxis = $chartXml.SelectSingleNode("//*[local-name()='catAx']")
        $categoryTitle = $categoryAxis.SelectSingleNode("*[local-name()='title']")
        $categoryTitle | Should -Not -BeNullOrEmpty
        $categoryTitle.InnerText | Should -Be 'Month'
        $categoryAxis.SelectSingleNode("*[local-name()='majorGridlines']") | Should -BeNullOrEmpty
        $categoryAxis.SelectSingleNode("*[local-name()='minorGridlines']") | Should -BeNullOrEmpty

        $valueAxis = $chartXml.SelectSingleNode("//*[local-name()='valAx']")
        $valueAxis | Should -Not -BeNullOrEmpty
        $valueAxis.SelectSingleNode("*[local-name()='title']").InnerText | Should -Be 'Revenue'
        $valueAxis.SelectSingleNode("*[local-name()='numFmt']").GetAttribute('formatCode') | Should -Be '$#,##0'
        $valueAxis.SelectSingleNode("*[local-name()='scaling']/*[local-name()='min']").GetAttribute('val') | Should -Be '0'
        $valueAxis.SelectSingleNode("*[local-name()='majorUnit']").GetAttribute('val') | Should -Be '100'
        $valueAxis.SelectSingleNode("*[local-name()='majorGridlines']") | Should -Not -BeNullOrEmpty
        $valueAxis.SelectSingleNode("*[local-name()='minorGridlines']") | Should -Not -BeNullOrEmpty

        $chartOuterXml | Should -Match 'trendline'
        $chartOuterXml | Should -Match 'dispEq'
        $chartOuterXml | Should -Match 'dispRSqr'
        $chartOuterXml | Should -Match '1F4E79'
        $chartOuterXml | Should -Match '4472C4'
        $chartOuterXml | Should -Match '70AD47'
        $chartOuterXml | Should -Match '7030A0'
        $chartOuterXml | Should -Match 'C00000'
    }

    It 'exports readable collection values and handles failing script properties' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelPowerShellProjection.xlsx'
        $row = [PSCustomObject]@{
            Name = 'Alpha'
            Tags = @('one', 'two')
        }
        $row | Add-Member -MemberType ScriptProperty -Name Broken -Value { throw 'boom' }
        $row | Export-OfficeExcel -Path $path -WorksheetName 'Data' -TableName 'Rows'

        $imported = @(Import-OfficeExcel -Path $path -WorksheetName 'Data')
        $imported.Count | Should -Be 1
        $imported[0].Name | Should -Be 'Alpha'
        $imported[0].Tags | Should -Be 'one, two'
        $imported[0].PSObject.Properties.Name | Should -Not -Contain 'Broken'

        $strictPath = Join-Path $TestDrive 'ExportOfficeExcelPowerShellProjectionStrict.xlsx'
        { $row | Export-OfficeExcel -Path $strictPath -PropertyConversionErrorAction Stop -ErrorAction Stop } |
            Should -Throw -ExpectedMessage "*Unable to read PowerShell property 'Broken'*"

        $placeholderPath = Join-Path $TestDrive 'ExportOfficeExcelPowerShellProjectionPlaceholder.xlsx'
        $row | Export-OfficeExcel -Path $placeholderPath -IncludeUnexportableProperties
        $placeholder = @(Import-OfficeExcel -Path $placeholderPath -WorksheetName 'Sheet1')
        $placeholder[0].Broken | Should -BeLike 'Property export failed:*boom*'

        if (-not ('PSWriteOffice.Tests.ExcelClrProjectionRow' -as [type])) {
            Add-Type -TypeDefinition @'
namespace PSWriteOffice.Tests {
    using System;

    public sealed class ExcelClrProjectionRow {
        public string Name { get { return "Alpha"; } }
        public string Broken { get { throw new InvalidOperationException("boom"); } }
        public string[] Tags { get { return new[] { "one", "two" }; } }
    }
}
'@
        }

        $clrRow = [PSWriteOffice.Tests.ExcelClrProjectionRow]::new()
        $clrPath = Join-Path $TestDrive 'ExportOfficeExcelClrProjection.xlsx'
        $clrRow | Export-OfficeExcel -Path $clrPath -WorksheetName 'Data' -TableName 'Rows'
        $clrImported = @(Import-OfficeExcel -Path $clrPath -WorksheetName 'Data')
        $clrImported[0].Name | Should -Be 'Alpha'
        $clrImported[0].Tags | Should -Be 'one, two'
        $clrImported[0].PSObject.Properties.Name | Should -Not -Contain 'Broken'

        $clrStrictPath = Join-Path $TestDrive 'ExportOfficeExcelClrProjectionStrict.xlsx'
        { $clrRow | Export-OfficeExcel -Path $clrStrictPath -PropertyConversionErrorAction Stop -ErrorAction Stop } |
            Should -Throw -ExpectedMessage "*Unable to read CLR property 'Broken'*"

        $clrPlaceholderPath = Join-Path $TestDrive 'ExportOfficeExcelClrProjectionPlaceholder.xlsx'
        $clrRow | Export-OfficeExcel -Path $clrPlaceholderPath -IncludeUnexportableProperties
        $clrPlaceholder = @(Import-OfficeExcel -Path $clrPlaceholderPath -WorksheetName 'Sheet1')
        $clrPlaceholder[0].Broken | Should -BeLike 'Property export failed:*boom*'
    }

    It 'applies strict property conversion to columns found only in later rows' {
        $path = Join-Path $TestDrive 'ExportOfficeExcelFirstRowProjection.xlsx'
        $first = [PSCustomObject]@{ Name = 'Alpha' }
        $second = [PSCustomObject]@{ Name = 'Beta' }
        $second | Add-Member -MemberType ScriptProperty -Name Broken -Value { throw 'later-only property was evaluated' }

        { @($first, $second) | Export-OfficeExcel -Path $path -WorksheetName 'Data' -PropertyConversionErrorAction Stop -ErrorAction Stop } |
            Should -Throw -ExpectedMessage "*Unable to read PowerShell property 'Broken'*"
    }

    It 'sets category date-axis scale values through the chart axis cmdlet' {
        $path = Join-Path $TestDrive 'DslExcelCategoryAxisScale.xlsx'
        $rows = @(
            [PSCustomObject]@{ Date = [datetime] '2026-01-01'; Revenue = 100 }
            [PSCustomObject]@{ Date = [datetime] '2026-01-15'; Revenue = 200 }
            [PSCustomObject]@{ Date = [datetime] '2026-02-01'; Revenue = 150 }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'Sales' -AutoFit
                $chart = Add-OfficeExcelChart -TableName 'Sales' -Row 6 -Column 1 -Type Line -Title 'Revenue Trend' -PassThru
                $chart |
                    Set-OfficeExcelChartAxis -CategoryNumberFormat 'yyyy-mm-dd' -CategoryMinimum 46000 -CategoryMaximum 46100 -CategoryMajorUnit 14 -CategoryMinorUnit 7 |
                    Out-Null
            }
        } | Out-Null

        $chartXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/drawings/charts/chart1.xml'
        $dateAxis = $chartXml.SelectSingleNode("//*[local-name()='dateAx']")
        $dateAxis | Should -Not -BeNullOrEmpty
        $dateAxis.SelectSingleNode("*[local-name()='scaling']/*[local-name()='min']").GetAttribute('val') | Should -Be '46000'
        $dateAxis.SelectSingleNode("*[local-name()='scaling']/*[local-name()='max']").GetAttribute('val') | Should -Be '46100'
        $dateAxis.SelectSingleNode("*[local-name()='majorUnit']").GetAttribute('val') | Should -Be '14'
        $dateAxis.SelectSingleNode("*[local-name()='minorUnit']").GetAttribute('val') | Should -Be '7'
    }

    It 'imports all sheets and can emit columns' {
        $path = Join-Path $TestDrive 'DslExcelImportAllSheetsByColumn.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Sales' -Content {
                Add-OfficeExcelTable -InputObject @(
                    [PSCustomObject]@{ Name = 'Alpha'; Value = 10; WorksheetName = 'SourceColumn' }
                ) -TableName 'SalesRows'
            }
            Add-OfficeExcelSheet -Name 'Inventory' -Content {
                Add-OfficeExcelTable -InputObject @(
                    [PSCustomObject]@{ Name = 'Widget'; Value = 4 }
                ) -TableName 'InventoryRows'
            }
        } | Out-Null

        $allRows = @(Import-OfficeExcel -Path $path -AllSheets)
        $allRows.Count | Should -Be 2
        @($allRows | Select-Object -ExpandProperty WorksheetName | Sort-Object) | Should -Be @('Inventory', 'Sales')
        ($allRows | Where-Object WorksheetName -eq 'Sales').WorksheetNameValue | Should -Be 'SourceColumn'

        $allRowsAsHashtable = @(Import-OfficeExcel -Path $path -AllSheets -AsHashtable)
        ($allRowsAsHashtable | Where-Object { $_['WorksheetName'] -eq 'Sales' })['WorksheetNameValue'] | Should -Be 'SourceColumn'

        $columns = @(Import-OfficeExcel -Path $path -WorksheetName 'Sales' -ByColumn)
        $columns.Count | Should -Be 3
        $columns[0].ColumnName | Should -Be 'Name'
        $columns[0].ColumnIndex | Should -Be 1
        @($columns[0].Values)[0] | Should -Be 'Alpha'
        $columns[1].ColumnName | Should -Be 'Value'
        @($columns[1].Values)[0] | Should -Be 10
    }

    It 'surfaces table style, number format, runtime, and write-reservation diagnostics' {
        $path = Join-Path $TestDrive 'DslExcelDiagnostics.xlsx'

        $styles = @(Get-OfficeExcelTableStyle -RecommendedOnly)
        $styles.Count | Should -BeGreaterThan 0
        ($styles | Where-Object Name -eq 'TableStyleMedium2').IsRecommended | Should -BeTrue

        $currency = Get-OfficeExcelNumberFormatPreset -CultureName en-US -Decimals 2 |
            Where-Object Name -eq 'Currency' |
            Select-Object -First 1
        $currency.FormatCode | Should -Be '"$"#,##0.00'

        $runtime = Get-OfficeExcelRuntimePreflight
        $runtime.FrameworkDescription | Should -Not -BeNullOrEmpty
        $runtime.PSObject.Properties.Name | Should -Contain 'Warnings'
        @($runtime.Warnings).Count | Should -BeGreaterOrEqual 0

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Ready'
            }
        } | Out-Null

        $reservation = Set-OfficeExcelWriteReservation -Path $path -ReadOnlyRecommended -UserName 'Reporting Team' -LegacyPasswordHash 'CAFE' -PassThru
        $reservation.Exists | Should -BeTrue
        $reservation.ReadOnlyRecommended | Should -BeTrue
        $reservation.UserName | Should -Be 'Reporting Team'
        $reservation.HasPasswordHash | Should -BeTrue

        $loaded = Get-OfficeExcelWriteReservation -Path $path
        $loaded.LegacyPasswordHash | Should -Be 'CAFE'

        $cleared = Clear-OfficeExcelWriteReservation -Path $path -PassThru
        $cleared.Exists | Should -BeFalse
    }

    It 'supports url images and smart hyperlink helpers' {
        $path = Join-Path $TestDrive 'DslExcelLinksAndImages.xlsx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive
        $imageUrl = [System.Uri]::new($imagePath).AbsoluteUri

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Reference'
                Set-OfficeExcelCell -Address 'B1' -Value 'Host'
                Set-OfficeExcelSmartHyperlink -Address 'A2' -Url 'https://datatracker.ietf.org/doc/html/rfc7208'
                Set-OfficeExcelHostHyperlink -Address 'B2' -Url 'https://learn.microsoft.com/office/open-xml/'
                Add-OfficeExcelImageFromUrl -Address 'D2' -Url $imageUrl -WidthPixels 32 -HeightPixels 32
                Add-OfficeExcelImage -Address 'E2' -Url $imageUrl -WidthPixels 24 -HeightPixels 24
            }
        } | Out-Null

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'xl/media/*' }).Count | Should -BeGreaterThan 0

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $smartText = $null
            $hostText = $null
            $doc['Data'].TryGetCellText(2, 1, [ref] $smartText) | Should -BeTrue
            $doc['Data'].TryGetCellText(2, 2, [ref] $hostText) | Should -BeTrue
            $smartText | Should -Be 'RFC 7208'
            $hostText | Should -Be 'learn.microsoft.com'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 2
    }

    It 'supports scaled and range-anchored worksheet images' {
        $path = Join-Path $TestDrive 'DslExcelImagePlacement.xlsx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Images' -Content {
                Add-OfficeExcelImage -Range 'A1:C3' -Path $imagePath -Name 'RangeLogo' -AltText 'Logo pinned to report header' -Title 'Pinned logo' -Placement MoveAndSize
                Add-OfficeExcelImage -Address 'E2' -Path $imagePath -ScalePercent 100 -Name 'ScaledLogo' -AltText 'Scaled logo' -RotationDegrees 12
            }
        } | Out-Null

        $drawingXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/drawings/drawing1.xml'
        $twoCellAnchor = $drawingXml.SelectSingleNode("/*[local-name()='wsDr']/*[local-name()='twoCellAnchor']")
        $twoCellAnchor | Should -Not -BeNullOrEmpty
        $twoCellAnchor.SelectSingleNode("*[local-name()='from']/*[local-name()='col']").InnerText | Should -Be '0'
        $twoCellAnchor.SelectSingleNode("*[local-name()='from']/*[local-name()='row']").InnerText | Should -Be '0'
        $twoCellAnchor.SelectSingleNode("*[local-name()='to']/*[local-name()='col']").InnerText | Should -Be '3'
        $twoCellAnchor.SelectSingleNode("*[local-name()='to']/*[local-name()='row']").InnerText | Should -Be '3'

        $rangeProperties = $twoCellAnchor.SelectSingleNode(".//*[local-name()='cNvPr']")
        $rangeProperties.GetAttribute('name') | Should -Be 'RangeLogo'
        $rangeProperties.GetAttribute('descr') | Should -Be 'Logo pinned to report header'
        $rangeProperties.GetAttribute('title') | Should -Be 'Pinned logo'

        $oneCellAnchor = $drawingXml.SelectSingleNode("/*[local-name()='wsDr']/*[local-name()='oneCellAnchor']")
        $oneCellAnchor | Should -Not -BeNullOrEmpty
        $scaledProperties = $oneCellAnchor.SelectSingleNode(".//*[local-name()='cNvPr']")
        $scaledProperties.GetAttribute('name') | Should -Be 'ScaledLogo'
        $scaledProperties.GetAttribute('descr') | Should -Be 'Scaled logo'
        $oneCellAnchor.SelectSingleNode(".//*[local-name()='xfrm']").GetAttribute('rot') | Should -Be '720000'
    }

    It 'supports internal link helpers for summary sheets' {
        $path = Join-Path $TestDrive 'DslExcelInternalLinks.xlsx'
        $rows = @(
            [PSCustomObject]@{ Sheet = 'Alpha'; Target = 'Alpha' }
            [PSCustomObject]@{ Sheet = 'Beta'; Target = 'Beta' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'SummaryTable' -AutoFit
                Set-OfficeExcelCell -Address 'D1' -Value 'Sheet'
                Set-OfficeExcelCell -Address 'D2' -Value 'Alpha'
                Set-OfficeExcelCell -Address 'D3' -Value 'Beta'
                Set-OfficeExcelInternalLinks -Range 'D2:D3'
                Set-OfficeExcelInternalLinksByHeader -Header 'Sheet' -TableName 'SummaryTable' -DisplayScript { param($text) "Open $text" }
            }
            Add-OfficeExcelSheet -Name 'Alpha' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Alpha Home'
            }
            Add-OfficeExcelSheet -Name 'Beta' -Content {
                Set-OfficeExcelCell -Address 'A1' -Value 'Beta Home'
            }
        } | Out-Null

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $summarySheet = $doc['Summary']
            $tableLink1 = $null
            $tableLink2 = $null
            $rangeLink1 = $null
            $rangeLink2 = $null
            $summarySheet.TryGetCellText(2, 1, [ref] $tableLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 1, [ref] $tableLink2) | Should -BeTrue
            $summarySheet.TryGetCellText(2, 4, [ref] $rangeLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 4, [ref] $rangeLink2) | Should -BeTrue
            $tableLink1 | Should -Be 'Open Alpha'
            $tableLink2 | Should -Be 'Open Beta'
            $rangeLink1 | Should -Be 'Alpha'
            $rangeLink2 | Should -Be 'Beta'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 4
    }

    It 'supports external URL link helpers for summary sheets' {
        $path = Join-Path $TestDrive 'DslExcelUrlLinks.xlsx'
        $rows = @(
            [PSCustomObject]@{ RFC = 'rfc7208'; Spec = 'rfc5321' }
            [PSCustomObject]@{ RFC = 'rfc7489'; Spec = 'rfc1035' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Summary' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'LinksTable' -AutoFit
                Set-OfficeExcelCell -Address 'D1' -Value 'Spec'
                Set-OfficeExcelCell -Address 'D2' -Value 'rfc5321'
                Set-OfficeExcelCell -Address 'D3' -Value 'rfc1035'

                Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -TableName 'LinksTable' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } -TitleScript { param($text) "Open $text" }
                Set-OfficeExcelUrlLinks -Range 'D2:D3' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" }
            }
        } | Out-Null

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $summarySheet = $doc['Summary']
            $tableLink1 = $null
            $tableLink2 = $null
            $rangeLink1 = $null
            $rangeLink2 = $null
            $summarySheet.TryGetCellText(2, 1, [ref] $tableLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 1, [ref] $tableLink2) | Should -BeTrue
            $summarySheet.TryGetCellText(2, 4, [ref] $rangeLink1) | Should -BeTrue
            $summarySheet.TryGetCellText(3, 4, [ref] $rangeLink2) | Should -BeTrue
            $tableLink1 | Should -Be 'Open rfc7208'
            $tableLink2 | Should -Be 'Open rfc7489'
            $rangeLink1 | Should -Be 'RFC 5321'
            $rangeLink2 | Should -Be 'RFC 1035'
        } finally {
            Close-OfficeExcel -Document $doc
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $hyperlinks = $sheetXml.SelectNodes("/*[local-name()='worksheet']/*[local-name()='hyperlinks']/*[local-name()='hyperlink']")
        $hyperlinks.Count | Should -Be 4
    }

    It 'styles Excel columns by header without range math' {
        $path = Join-Path $TestDrive 'DslExcelColumnStyleByHeader.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Revenue = 1200.5; Rate = 0.42; Status = 'Ready' }
            [PSCustomObject]@{ Name = 'Beta'; Revenue = 800.25; Rate = 0.18; Status = 'Blocked' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'ReportRows'
                Set-OfficeExcelColumnStyleByHeader -Header Revenue -Style Currency -CultureName en-US -AutoFit
                Set-OfficeExcelColumnStyleByHeader -Header Rate -Style Percent -Decimals 1
                Set-OfficeExcelColumnStyleByHeader -Header Status -BackgroundByText @{ Ready = '#D4EDDA'; Blocked = '#F8D7DA' } -BoldByText Blocked
            }
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $revenueCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='B2']")
        $rateCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='C2']")
        $statusCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='D3']")

        $revenueCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $rateCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $statusCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
    }

    It 'preserves case-distinct text style map entries when requested' {
        $path = Join-Path $TestDrive 'DslExcelColumnStyleByHeaderCaseSensitive.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Status = 'Ready' }
            [PSCustomObject]@{ Name = 'Beta'; Status = 'ready' }
        )
        $statusColors = [hashtable]::new([System.StringComparer]::Ordinal)
        $statusColors.Add('Ready', '#D4EDDA')
        $statusColors.Add('ready', '#F8D7DA')

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Data' -Content {
                Add-OfficeExcelTable -InputObject $rows -TableName 'ReportRows'
                Set-OfficeExcelColumnStyleByHeader -Header Status -BackgroundByText $statusColors -CaseSensitive
            }
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $upperCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='B2']")
        $lowerCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='B3']")

        $upperCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $lowerCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $upperCell.GetAttribute('s') | Should -Not -Be $lowerCell.GetAttribute('s')
    }

    It 'creates composed Excel report sheets from PowerShell blocks' {
        $path = Join-Path $TestDrive 'DslExcelReportSheet.xlsx'
        $rows = @(
            [PSCustomObject]@{ Name = 'Alpha'; Score = 9; Status = 'Ready' }
            [PSCustomObject]@{ Name = 'Beta'; Score = 4; Status = 'Blocked' }
        )

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Summary' {
                Add-OfficeExcelReportTitle -Title 'Operational Summary' -Subtitle 'Current view'
                Add-OfficeExcelReportKpiRow -Data ([ordered] @{ Ready = 1; Blocked = 1 }) -PerRow 2
                Add-OfficeExcelReportCallout -Kind Warning -Title 'Attention' -Body 'One item needs review.'
                Add-OfficeExcelReportTable -Data $rows -Title 'Rows'
                Add-OfficeExcelReportLegend -Title 'Legend' -Headers 'Status','Meaning' -Rows @(
                    @('Ready', 'No action'),
                    @('Blocked', 'Needs owner')
                ) -FirstColumnFillByValue @{ Ready = '#D4EDDA'; Blocked = '#F8D7DA' }
            }
        }

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $sheet = $doc['Summary']
            $title = $null
            $subtitle = $null
            $readyLabel = $null
            $calloutTitle = $null

            $sheet.TryGetCellText(1, 1, [ref] $title) | Should -BeTrue
            $sheet.TryGetCellText(2, 1, [ref] $subtitle) | Should -BeTrue
            $sheet.TryGetCellText(4, 1, [ref] $readyLabel) | Should -BeTrue
            $sheet.TryGetCellText(7, 1, [ref] $calloutTitle) | Should -BeTrue

            $title | Should -Be 'Operational Summary'
            $subtitle | Should -Be 'Current view'
            $readyLabel | Should -Be 'Ready'
            $calloutTitle | Should -Be 'Attention'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'preserves case-distinct report legend fill entries when requested' {
        $path = Join-Path $TestDrive 'DslExcelReportLegendCaseSensitive.xlsx'
        $statusColors = [hashtable]::new([System.StringComparer]::Ordinal)
        $statusColors.Add('Ready', '#D4EDDA')
        $statusColors.Add('ready', '#F8D7DA')

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Legend' {
                Add-OfficeExcelReportLegend -Header 'Status','Meaning' -InputObject @(
                    @('Ready', 'Upper'),
                    @('ready', 'Lower')
                ) -FirstColumnFillByValue $statusColors -CaseSensitive
            }
        }

        $sheetXml = Get-ZipXmlDocumentLocal -Path $path -Entry 'xl/worksheets/sheet1.xml'
        $upperCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='A2']")
        $lowerCell = $sheetXml.SelectSingleNode("//*[local-name()='c' and @r='A3']")

        $upperCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $lowerCell.GetAttribute('s') | Should -Not -BeNullOrEmpty
        $upperCell.GetAttribute('s') | Should -Not -Be $lowerCell.GetAttribute('s')
    }

    It 'uses the topmost report composer for nested report sheets' {
        $path = Join-Path $TestDrive 'DslExcelNestedReportSheets.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelReportSheet -Name 'Outer' {
                Add-OfficeExcelReportTitle -Title 'Outer title'
                Add-OfficeExcelReportSheet -Name 'Inner' {
                    Add-OfficeExcelReportTitle -Title 'Inner title'
                }
                Add-OfficeExcelReportParagraph -Text 'Outer after inner'
            }
        }

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $outer = $doc['Outer']
            $inner = $doc['Inner']
            $outerTitle = $null
            $outerAfter = $null
            $innerTitle = $null

            $outer.TryGetCellText(1, 1, [ref] $outerTitle) | Should -BeTrue
            $outer.TryGetCellText(3, 1, [ref] $outerAfter) | Should -BeTrue
            $inner.TryGetCellText(1, 1, [ref] $innerTitle) | Should -BeTrue

            $outerTitle | Should -Be 'Outer title'
            $outerAfter | Should -Be 'Outer after inner'
            $innerTitle | Should -Be 'Inner title'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'uses the topmost worksheet for nested sheet blocks' {
        $path = Join-Path $TestDrive 'DslExcelNestedSheets.xlsx'

        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Outer' {
                Set-OfficeExcelCell -Address 'A1' -Value 'Outer start'
                Add-OfficeExcelSheet -Name 'Inner' {
                    Set-OfficeExcelCell -Address 'A1' -Value 'Inner value'
                }
                Set-OfficeExcelCell -Address 'A2' -Value 'Outer after inner'
            }
        }

        $doc = Get-OfficeExcel -Path $path -ReadOnly
        try {
            $outer = $doc['Outer']
            $inner = $doc['Inner']
            $outerStart = $null
            $outerAfter = $null
            $innerValue = $null

            $outer.TryGetCellText(1, 1, [ref] $outerStart) | Should -BeTrue
            $outer.TryGetCellText(2, 1, [ref] $outerAfter) | Should -BeTrue
            $inner.TryGetCellText(1, 1, [ref] $innerValue) | Should -BeTrue

            $outerStart | Should -Be 'Outer start'
            $outerAfter | Should -Be 'Outer after inner'
            $innerValue | Should -Be 'Inner value'
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'surfaces workbook intelligence workflows through thin commands' {
        $left = Join-Path $TestDrive 'DslExcelWorkbookIntelligenceLeft.xlsx'
        $right = Join-Path $TestDrive 'DslExcelWorkbookIntelligenceRight.xlsx'

        New-OfficeExcel -Path $left {
            Add-OfficeExcelSheet -Name 'Data' {
                Set-OfficeExcelCell -Address A1 -Value '{{CustomerName}}'
                Set-OfficeExcelCell -Address B1 -Value 20
                Set-OfficeExcelFormula -Address C1 -Formula 'SUM(A1:B1)+NOW()'
                Set-OfficeExcelNamedRange -Name 'Totals' -Range 'A1:C1'
                Add-OfficeExcelComment -Address A1 -Text 'Needs binding review' -Author 'Reviewer'
            }
        }

        New-OfficeExcel -Path $right {
            Add-OfficeExcelSheet -Name 'Data' {
                Set-OfficeExcelCell -Address A1 -Value 11
                Set-OfficeExcelCell -Address B1 -Value 20
                Set-OfficeExcelFormula -Address C1 -Formula 'SUM(A1:B1)'
            }
        }

        $doctor = Test-OfficeExcelWorkbook -Path $left -SkipOpenXmlValidation
        $doctor.Passed | Should -BeTrue

        $formulas = Get-OfficeExcelFormulaAnalysis -Path $left -IncludeFormulas
        $formulas.FormulaCount | Should -Be 1
        $formulas.VolatileFormulaCount | Should -Be 1
        $formulas.Formulas[0].Functions | Should -Contain 'SUM'

        $streaming = Get-OfficeExcelStreamingContract -Path $left
        $streaming.WorksheetCount | Should -Be 1

        $dataModel = Get-OfficeExcelDataModel -Path $left
        $dataModel.HasDataModelOrQueries | Should -BeFalse

        $accessibility = Test-OfficeExcelAccessibility -Path $left
        $accessibility.Passed | Should -BeTrue

        $diff = Compare-OfficeExcelWorkbook -Path $left -DifferencePath $right
        $diff.AreEqual | Should -BeFalse
        $diff.DifferenceCount | Should -BeGreaterThan 0
        $diff.Differences.Category | Should -Contain 'Comment'

        $commentAudit = Get-OfficeExcelCommentAudit -Path $left -IncludeComments
        $commentAudit.CommentCount | Should -Be 1
        $commentAudit.Comments[0].Author | Should -Be 'Reviewer'

        $threaded = Add-OfficeExcelThreadedComment -Path $left -Sheet Data -Address A1 -Text 'Threaded review note' -Author 'Modern Reviewer' -PassThru
        $threaded.CellReference | Should -Be 'A1'
        $threaded.Author | Should -Be 'Modern Reviewer'
        $threadedReply = Add-OfficeExcelThreadedComment -Path $left -Sheet Data -Address A1 -Text 'Threaded reply' -Author 'Report Owner' -ParentId $threaded.Id -Done -PassThru
        $threadedReply.IsReply | Should -BeTrue
        $threadedReply.Done | Should -BeTrue
        $commentAuditAfterThreaded = Get-OfficeExcelCommentAudit -Path $left -IncludeComments
        $commentAuditAfterThreaded.ThreadedCommentCount | Should -Be 2
        $commentAuditAfterThreaded.ThreadedComments.Author | Should -Contain 'Modern Reviewer'
        $commentAuditAfterThreaded.ThreadedComments.SheetName | Should -Contain 'Data'

        $template = Test-OfficeExcelTemplateBinding -Path $left -Binding @{ CustomerName = 'Northwind' }
        $template.Passed | Should -BeTrue
        (Test-OfficeExcelTemplateBinding -Path $left -Binding @{} -Quiet) | Should -BeFalse

        $metadata = Add-OfficeExcelPowerQueryMetadata -Path $left -Name 'CustomerQuery' -WorksheetName 'Data' -CommandText 'let Source = Excel.CurrentWorkbook() in Source' -RefreshOnOpen -PassThru
        $metadata.AddedWorkbookConnection | Should -BeTrue
        $metadata.AddedWorksheetQueryTable | Should -BeTrue
        $metadata.ConnectionId | Should -Be 1
        $metadata.QueryTableName | Should -Be 'CustomerQueryTable'
        $metadata2 = Add-OfficeExcelPowerQueryMetadata -Path $left -Name 'CustomerQuery2' -WorksheetName 'Data' -CommandText 'let Source = Excel.CurrentWorkbook() in Source' -PassThru
        $metadata2.ConnectionId | Should -Be 2
        $dataModelAfterMetadata = Get-OfficeExcelDataModel -Path $left
        $dataModelAfterMetadata.HasDataModelOrQueries | Should -BeTrue

        $repair = Repair-OfficeExcelWorkbook -Path $left -PassThru
        $repair.ActionCount | Should -BeGreaterThan 0

        $doc = Get-OfficeExcel -Path $left
        try {
            ($doc | Rename-OfficeExcelNamedRange -Name 'Totals' -NewName 'GrandTotal' -Sheet 'Data' -PassThru) | Should -BeTrue
            ($doc | Remove-OfficeExcelNamedRange -Name 'GrandTotal' -Sheet 'Data' -PassThru) | Should -BeTrue
            Save-OfficeExcel -Document $doc
        } finally {
            Close-OfficeExcel -Document $doc
        }
    }

    It 'imports normalized delimited text into Excel' {
        $path = Join-Path $TestDrive 'DslExcelDelimitedImport.xlsx'
        $csv = Join-Path $TestDrive 'DelimitedImport.csv'

        Set-Content -Path $csv -Value "Name;Amount`r`nAlpha;10.5`r`nBeta;11.75" -NoNewline
        New-OfficeExcel -Path $path {
            Add-OfficeExcelSheet -Name 'Seed'
        }

        $result = Import-OfficeExcelDelimitedText -Path $path -SourcePath $csv -Delimiter ';' -SheetName 'Import' -PassThru
        $result.SheetName | Should -Be 'Import'
        $result.RowCount | Should -Be 2
        $result.ColumnCount | Should -Be 2

        $rows = @(Import-OfficeExcel -Path $path -WorksheetName 'Import' -Range 'A1:B3')
        $rows[0].Name | Should -Be 'Alpha'
        $rows[1].Amount | Should -Be 11.75
    }
}
