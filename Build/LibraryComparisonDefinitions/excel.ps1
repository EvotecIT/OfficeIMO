# Excel benchmark suites in OfficeIMO.Excel.Benchmarks beyond the 65K typed-read and data-reader write comparisons
# declared in Build/LibraryComparisonDefinitions.ps1. Each workload selects one BenchmarkDotNet class.
# Returns an ordered dictionary of workloads; loading this file performs no work.
param(
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0'
)

$project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
$platforms = @('windows', 'linux', 'macos')

# Every method x parameter combination, with parameters given as ordered name/value-list pairs.
function New-ExcelExpectedCases {
    param(
        [Parameter(Mandatory)] [string[]] $Methods,
        [System.Collections.Specialized.OrderedDictionary] $Parameters = ([ordered]@{})
    )

    $combinations = @('')
    foreach ($parameter in $Parameters.GetEnumerator()) {
        $combinations = @(
            foreach ($prefix in $combinations) {
                foreach ($value in @($parameter.Value)) {
                    $pair = '{0}={1}' -f $parameter.Key, $value
                    if ($prefix) { "$prefix&$pair" } else { $pair }
                }
            }
        )
    }
    foreach ($method in $Methods) {
        foreach ($combination in $combinations) {
            if ($combination) { "$method|$combination" } else { $method }
        }
    }
}

$formats = @('Xlsx', 'Xlsb', 'Xls')
$booleans = @('False', 'True')

[ordered]@{
    # Library comparisons ----------------------------------------------------------------------------------------

    excelread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX sales report reading into objects and DataTables'
        Libraries = @('OfficeIMO', 'ClosedXML')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelReadBenchmarks.*'
        ComparisonId = 'excel-xlsx-sales-read'
        Suite = 'OfficeIMO.Excel.SalesRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @(
                'OfficeIMO_Read_Objects', 'OfficeIMO_Read_DataTable', 'OfficeIMO_Read_DataTableCount',
                'ClosedXML_Read_DataTableCount', 'ClosedXML_Read_Objects', 'ClosedXML_Read_DataTable') -Parameters ([ordered]@{
                    RowCount = @(250, 2500, 25000)
                }))
    }
    excelwrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX sales report writing'
        Libraries = @('OfficeIMO', 'ClosedXML')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelWriteBenchmarks.*'
        ComparisonId = 'excel-xlsx-sales-report-write'
        Suite = 'OfficeIMO.Excel.SalesReportWrite'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_Write_Report', 'ClosedXML_Write_Report') -Parameters ([ordered]@{
                    RowCount = @(250, 2500, 25000, 100000)
                }))
    }
    excelroundtrip = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX load, edit and save'
        Libraries = @('OfficeIMO', 'ClosedXML')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelRoundTripBenchmarks.*'
        ComparisonId = 'excel-xlsx-load-edit-save'
        Suite = 'OfficeIMO.Excel.LoadEditSave'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_Load_Edit_Save', 'ClosedXML_Load_Edit_Save') -Parameters ([ordered]@{
                    RowCount = @(250, 2500, 25000)
                }))
    }
    excelproducerroundtrip = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX open, edit and save of workbooks from other producers'
        Libraries = @('OfficeIMO', 'ClosedXML')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelProducerRoundTripBenchmarks.*'
        ComparisonId = 'excel-xlsx-producer-edit'
        Suite = 'OfficeIMO.Excel.ProducerEdit'
        # OFFICEIMO_EXCEL_PRODUCER_FIXTURES adds a Producer value per workbook; declared runs leave it unset.
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_Open_Edit_Save', 'ClosedXML_Open_Edit_Save') -Parameters ([ordered]@{
                    Producer = @('OfficeIMO', 'ClosedXML', 'Excel')
                }))
    }
    excelstyleheavywrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX writing with 240 distinct cell styles'
        Libraries = @('OfficeIMO', 'ClosedXML')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelStyleHeavyWriteBenchmarks.*'
        ComparisonId = 'excel-xlsx-style-heavy-write'
        Suite = 'OfficeIMO.Excel.StyleHeavyWrite'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_Write_Styled', 'ClosedXML_Write_Styled') -Parameters ([ordered]@{
                    RowCount = @(20000)
                }))
    }
    excelsmallworkbookloop = [pscustomobject]@{
        Area = 'excel'
        Title = 'Opening and reading 200 small XLSX workbooks'
        Libraries = @('OfficeIMO', 'ClosedXML', 'Open XML SDK')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelSmallWorkbookLoopBenchmarks.*'
        ComparisonId = 'excel-xlsx-small-workbook-loop'
        Suite = 'OfficeIMO.Excel.SmallWorkbookLoop'
        ExpectedCases = @('OfficeIMO_Open_Read', 'ClosedXML_Open_Read', 'OpenXml_Open_Read')
    }
    excelloadedvalueread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX read of every value, 50,000 x 15'
        Libraries = @('OfficeIMO', 'Open XML SDK')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelLoadedValueReadBenchmarks.*'
        ComparisonId = 'excel-xlsx-read-all-values'
        Suite = 'OfficeIMO.Excel.ReadAllValues'
        ExpectedCases = @('OfficeIMO', 'OpenXml')
    }
    excelwideloadedread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX wide-sheet presence scan, 10,000 x 150'
        Libraries = @('OfficeIMO', 'Open XML SDK')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'Supplemental scan outside the published read shapes; each lane is checked only by its count of non-empty cells, not by the values.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelWideLoadedReadBenchmarks.*'
        ComparisonId = 'excel-xlsx-wide-presence-scan'
        Suite = 'OfficeIMO.Excel.WidePresenceScan'
        ExpectedCases = @('OfficeIMO', 'OpenXml')
    }
    excelsharedstringread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX shared-string reading'
        Libraries = @('OfficeIMO', 'Sylvan.Data.Excel')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelSharedStringReadBenchmarks.*'
        ComparisonId = 'excel-xlsx-shared-string-read'
        Suite = 'OfficeIMO.Excel.SharedStringRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO', 'Sylvan') -Parameters ([ordered]@{
                    RowCount = @(256, 25000)
                    PrefixedWorksheet = $booleans
                    Shape = @('Ascii', 'UnicodeTail', 'EntityTail', 'Unicode', 'RichText')
                }))
    }
    exceltextwrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX export of short and long text'
        Libraries = @('OfficeIMO', 'SpreadCheetah')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelTextWriteBenchmarks.*'
        ComparisonId = 'excel-xlsx-text-write'
        Suite = 'OfficeIMO.Excel.TextWrite'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO', 'SpreadCheetah') -Parameters ([ordered]@{
                    TextLength = @(64, 4096)
                    TextShape = @('Plain', 'Escaped', 'Markup')
                }))
    }
    excellargetypedread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX typed reading up to a million rows'
        Libraries = @('OfficeIMO', 'Sylvan.Data.Excel', 'ExcelReader.NET')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelLargeTypedReadBenchmarks.*'
        ComparisonId = 'excel-xlsx-large-typed-read'
        Suite = 'OfficeIMO.Excel.LargeTypedRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO', 'Sylvan', 'ExcelReaderNet') -Parameters ([ordered]@{
                    RowCount = @(25000, 250000, 1000000)
                }))
    }
    excelrowstreaming = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX streaming write of a million generated rows'
        Libraries = @('OfficeIMO', 'ExcelReader.NET')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelGeneratedRowStreamingBenchmarks.*'
        ComparisonId = 'excel-xlsx-generated-row-streaming'
        Suite = 'OfficeIMO.Excel.GeneratedRowStreaming'
        MethodIdentity = @{
            'WriteRowsGenerated' = 'OfficeIMO|WriteRowsGenerated'
            'WriteRowsAsyncGenerated' = 'OfficeIMO|WriteRowsAsyncGenerated'
        }
        ExpectedCases = @(New-ExcelExpectedCases -Methods @(
                'WriteRowsGenerated', 'ExcelReaderNetWriteRowsGenerated',
                'WriteRowsAsyncGenerated', 'ExcelReaderNetWriteRowsAsyncGenerated') -Parameters ([ordered]@{
                    RowCount = @(1000000)
                }))
    }
    excelarrow = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX 65K conversion to Apache Arrow'
        Libraries = @('OfficeIMO', 'ExcelReader.NET')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Filter = '*.ExcelArrowConversionBenchmarks.*'
        ComparisonId = 'excel-xlsx-arrow-conversion'
        Suite = 'OfficeIMO.Excel.ArrowConversion'
        ExpectedCases = @('OfficeIMO_ExplicitSchema', 'ExcelReaderNet_ExplicitSchema', 'OfficeIMO_BoundedStreaming')
    }
    excelcsvtoxlsx = [pscustomobject]@{
        Area = 'excel'
        Title = 'CSV to XLSX conversion'
        Libraries = @('OfficeIMO', 'MiniExcel', 'Sylvan.Data.Excel')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.CsvToXlsxConversionBenchmarks.*'
        ComparisonId = 'excel-csv-to-xlsx'
        Suite = 'OfficeIMO.Excel.CsvToXlsx'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_ImportCsv', 'OfficeIMO_StreamingText', 'MiniExcel', 'Sylvan') -Parameters ([ordered]@{
                    Rows = @(50000)
                }))
    }
    excelxlsxtocsv = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX to CSV conversion'
        Libraries = @('OfficeIMO', 'MiniExcel', 'Sylvan.Data.Excel')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.XlsxToCsvConversionBenchmarks.*'
        ComparisonId = 'excel-xlsx-to-csv'
        Suite = 'OfficeIMO.Excel.XlsxToCsv'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_SaveAsCsv', 'OfficeIMO_Reader', 'MiniExcel', 'Sylvan') -Parameters ([ordered]@{
                    Rows = @(50000)
                }))
    }
    excelmultisheetread = [pscustomobject]@{
        Area = 'excel'
        Title = 'Multi-sheet workbook streaming read'
        Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sylvan.Data.Excel', 'ExcelDataReader')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelMultiSheetReadBenchmarks.*'
        ComparisonId = 'excel-multi-sheet-stream-read'
        Suite = 'OfficeIMO.Excel.MultiSheetStreamRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @(
                'OfficeIMO_Sequential', 'OfficeIMO_Parallel', 'ExcelReaderNet_Sequential', 'ExcelReaderNet_Parallel',
                'Sylvan_Sequential', 'ExcelDataReader_Sequential') -Parameters ([ordered]@{
                    Format = $formats
                    SheetCount = @(8)
                    RowsPerSheet = @(25000)
                }))
    }
    excelmultisheetworkbookread = [pscustomobject]@{
        Area = 'excel'
        Title = 'Multi-sheet XLSX workbook load and read'
        Libraries = @('OfficeIMO', 'ClosedXML')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelMultiSheetWorkbookReadBenchmarks.*'
        ComparisonId = 'excel-multi-sheet-workbook-read'
        Suite = 'OfficeIMO.Excel.MultiSheetWorkbookRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_Sequential', 'ClosedXml_Sequential') -Parameters ([ordered]@{
                    SheetCount = @(8)
                    RowsPerSheet = @(25000)
                }))
    }
    excelmultisheetwrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'Multi-sheet XLSX writing'
        Libraries = @('OfficeIMO', 'SpreadCheetah', 'ExcelReader.NET', 'MiniExcel', 'LargeXlsx')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelMultiSheetWriteBenchmarks.*'
        ComparisonId = 'excel-multi-sheet-write'
        Suite = 'OfficeIMO.Excel.MultiSheetWrite'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @(
                'OfficeIMO_Typed', 'OfficeIMO_TypedNoTable', 'OfficeIMO_Editable', 'SpreadCheetah', 'ExcelReaderNet',
                'MiniExcel', 'LargeXlsx') -Parameters ([ordered]@{
                    SheetCount = @(10, 50)
                    RowsPerSheet = @(5000)
                }))
    }
    excelorderedparallel = [pscustomobject]@{
        Area = 'excel'
        Title = 'Ordered parallel typed reading'
        Libraries = @('OfficeIMO', 'ExcelReader.NET')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelOrderedParallelComparisonBenchmarks.*'
        ComparisonId = 'excel-ordered-parallel-read'
        Suite = 'OfficeIMO.Excel.OrderedParallelRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_Ordered', 'ExcelReaderNet_Ordered') -Parameters ([ordered]@{
                    Format = @('Xlsx', 'Xlsb')
                    RequestedDop = @(1, 4)
                    EscapedNames = $booleans
                    InferSchema = $booleans
                }))
    }
    excelgeneratedread = [pscustomobject]@{
        Area = 'excel'
        Title = 'Generated four-column workbook reading and object mapping'
        Libraries = @('OfficeIMO', 'Sylvan.Data.Excel', 'ExcelDataReader', 'ExcelReader.NET')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelReaderGeneratedBenchmarks.*'
        ComparisonId = 'excel-generated-read'
        Suite = 'OfficeIMO.Excel.GeneratedRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @(
                'OfficeIMO_Read', 'Sylvan_Read', 'ExcelDataReader_Read', 'ExcelReaderNet_Read', 'OfficeIMO_Parse', 'ExcelReaderNet_Parse') -Parameters ([ordered]@{
                    Format = $formats
                    RowCount = @(50000)
                    EscapedNames = $booleans
                }))
    }
    excelzipprefetch = [pscustomobject]@{
        Area = 'excel'
        Title = 'Generated workbook reading with ZIP decompression prefetch'
        Libraries = @('OfficeIMO', 'ExcelReader.NET')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelReaderGeneratedZipPrefetchBenchmarks.*'
        ComparisonId = 'excel-generated-zip-prefetch-read'
        Suite = 'OfficeIMO.Excel.GeneratedZipPrefetchRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO', 'ExcelReaderNet', 'ExcelReaderNet_Prefetch') -Parameters ([ordered]@{
                    Format = @('Xlsx', 'Xlsb')
                    RowCount = @(50000)
                    EscapedNames = $booleans
                }))
    }
    excelgeneratedwrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'Generated four-column workbook writing'
        Libraries = @('OfficeIMO', 'ExcelReader.NET')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelReaderGeneratedWriteBenchmarks.*'
        ComparisonId = 'excel-generated-write'
        Suite = 'OfficeIMO.Excel.GeneratedWrite'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_WriteData', 'OfficeIMO_MatchingWriteData', 'OfficeIMO', 'ExcelReaderNet') -Parameters ([ordered]@{
                    Format = $formats
                    RowCount = @(50000)
                }))
    }
    excelgeneratedxlsxwrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'Generated four-column XLSX writing'
        Libraries = @('OfficeIMO', 'ExcelReader.NET', 'SpreadCheetah')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelReaderGeneratedXlsxWriteBenchmarks.*'
        ComparisonId = 'excel-generated-xlsx-write'
        Suite = 'OfficeIMO.Excel.GeneratedXlsxWrite'
        ExpectedCases = @('OfficeIMO', 'ExcelReaderNet', 'SpreadCheetah')
    }
    excelwriteroptions = [pscustomobject]@{
        Area = 'excel'
        Title = 'Generated workbook writing with serial, prefetch and parallel options'
        Libraries = @('OfficeIMO', 'ExcelReader.NET')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelWriterOptionsComparisonBenchmarks.*'
        ComparisonId = 'excel-generated-write-options'
        Suite = 'OfficeIMO.Excel.GeneratedWriteOptions'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @(
                'OfficeIMO_Serial', 'ExcelReaderNet_Serial', 'ExcelReaderNet_Prefetch', 'ExcelReaderNet_Parallel1', 'ExcelReaderNet_Parallel8') -Parameters ([ordered]@{
                    Format = @('Xlsx', 'Xlsb')
                    RowCount = @(50000)
                }))
    }
    exceltypedmapping = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX 65K typed object and DataTable mapping'
        Libraries = @('OfficeIMO', 'Sylvan.Data.Excel', 'MiniExcel', 'ExcelReader.NET', 'ExcelDataReader')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'Lanes are not validated yet: setup checks every lane against the CSV fixture but does not record the validation results.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.MarkPflug65KTypedMappingXlsxBenchmarks.*'
        ComparisonId = 'markpflug-65k-xlsx-typed-mapping'
        Suite = 'OfficeIMO.Excel.Xlsx.MarkPflug65KTypedMapping'
        ExpectedCases = @(
            'OfficeIMO_Class'
            'OfficeIMO_ClassParallel'
            'Sylvan_Class'
            'MiniExcel_Class'
            # The ExcelReader.NET parser lane is compiled only for the modern ExcelReader.NET API.
            if ($Framework -eq 'net10.0') { 'ExcelReaderNet_Class' }
            'OfficeIMO_Record'
            'OfficeIMO_DataTable'
            'ExcelDataReader_DataSet'
            'Sylvan_DataTable'
        )
    }

    # OfficeIMO-only suites ---------------------------------------------------------------------------------------

    excelbufferedtypedread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX buffered typed reading by execution mode'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only execution-mode and first-row variants; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelBufferedTypedReadBenchmarks.*'
        ComparisonId = 'excel-buffered-typed-read'
        Suite = 'OfficeIMO.Excel.BufferedTypedRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('Read', 'ReadAutomatic', 'ReadSequential', 'ReadParallel', 'FirstRow') -Parameters ([ordered]@{
                    RowCount = @(1000, 5000)
                    ColumnCount = @(4, 65)
                    ReverseRows = $booleans
                    Utf16 = $booleans
                }))
    }
    excelbufferedxlsbwrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSB buffered writing'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only XLSB writer profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelBufferedXlsbWriteBenchmarks.*'
        ComparisonId = 'excel-buffered-xlsb-write'
        Suite = 'OfficeIMO.Excel.BufferedXlsbWrite'
        ExpectedCases = @('Write|RowCount=50000')
    }
    excelcharacterread = [pscustomobject]@{
        Area = 'excel'
        Title = 'Spreadsheet string and character-chunk reading'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only string access variants; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelCharacterReadBenchmarks.*'
        ComparisonId = 'excel-character-read'
        Suite = 'OfficeIMO.Excel.CharacterRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('StringValues', 'CharacterChunks') -Parameters ([ordered]@{
                    Format = $formats
                    Shape = @('Short', 'Long')
                }))
    }
    exceldatatableexecution = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX DataTable loading by execution mode'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only execution-mode variants; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelDataTableExecutionBenchmarks.*'
        ComparisonId = 'excel-datatable-execution'
        Suite = 'OfficeIMO.Excel.DataTableExecution'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('Automatic', 'Sequential', 'Parallel') -Parameters ([ordered]@{
                    RowCount = @(25000)
                }))
    }
    exceldatatablerowshape = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX DataTable loading by row shape'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only row-shape profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelDataTableRowShapeBenchmarks.*'
        ComparisonId = 'excel-datatable-row-shape'
        Suite = 'OfficeIMO.Excel.DataTableRowShape'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('Read') -Parameters ([ordered]@{
                    RowCount = @(1000)
                    ColumnCount = @(8, 65)
                    Sparse = $booleans
                    RepeatedRow = $booleans
                    ReverseRows = $booleans
                }))
    }
    exceldefaulttextwrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX text export on default options'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only default-options profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelDefaultTextWriteBenchmarks.*'
        ComparisonId = 'excel-default-text-write'
        Suite = 'OfficeIMO.Excel.DefaultTextWrite'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO') -Parameters ([ordered]@{
                    TextLength = @(64, 4096)
                    TextShape = @('Plain', 'Escaped', 'Markup')
                }))
    }
    excelformattedreport = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX formatted report writing'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only formatted report profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelFormattedReportBenchmarks.*'
        ComparisonId = 'excel-formatted-report-write'
        Suite = 'OfficeIMO.Excel.FormattedReportWrite'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('WriteFormattedReport', 'WriteValidatedFormattedReport') -Parameters ([ordered]@{
                    RowCount = @(5000, 25000)
                    StyleInBatch = $booleans
                }))
    }
    excellargetypedfirstrow = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX first-row access up to a million rows'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only first-row contract; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelLargeTypedFirstRowBenchmarks.*'
        ComparisonId = 'excel-large-typed-first-row'
        Suite = 'OfficeIMO.Excel.LargeTypedFirstRow'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO') -Parameters ([ordered]@{
                    RowCount = @(25000, 250000, 1000000)
                }))
    }
    excellegacyxlsstringread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLS string record reading'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only legacy string-record profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelLegacyXlsStringReadBenchmarks.*'
        ComparisonId = 'excel-xls-string-read'
        Suite = 'OfficeIMO.Excel.XlsStringRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('Read') -Parameters ([ordered]@{
                    RowCount = @(5000)
                    Cells = @('Label', 'LabelUtf16', 'FormulaString', 'ContinuedSharedString')
                    Access = @('GetString', 'GetChars')
                }))
    }
    excelnativebinarymodelread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLS and XLSB workbook model loading'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only native binary model profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelNativeBinaryModelReadBenchmarks.*'
        ComparisonId = 'excel-native-binary-model-read'
        Suite = 'OfficeIMO.Excel.NativeBinaryModelRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('LoadModel') -Parameters ([ordered]@{
                    Format = @('Xls', 'Xlsb')
                    RowCount = @(20000)
                }))
    }
    excelnativebinarywrite = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLS and XLSB native writing'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only native binary writer profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.OfficeNativeBinaryWriteBenchmarks.*'
        ComparisonId = 'excel-native-binary-write'
        Suite = 'OfficeIMO.Excel.NativeBinaryWrite'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OfficeIMO_ValidatedNativeBinaryWrite') -Parameters ([ordered]@{
                    RowCount = @(2500, 25000)
                    Format = @('Xls', 'Xlsb')
                }))
    }
    excelnumericxmlread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX numeric worksheet reading by API'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only API and worksheet-markup profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelNumericXmlReadBenchmarks.*'
        ComparisonId = 'excel-numeric-xml-read'
        Suite = 'OfficeIMO.Excel.NumericXmlRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('Read') -Parameters ([ordered]@{
                    RowCount = @(2500, 25000)
                    ReadRowCount = @(0)
                    NumericAsDecimal = $booleans
                    Utf16 = @('True')
                    EnableWorksheetPrefetch = @('False')
                    InferDataTableColumnTypes = $booleans
                    Api = @('DataReader', 'TypedDataReader', 'Range', 'RangeStream', 'UsedRange', 'DataTable', 'Objects', 'ObjectsStream')
                    Coordinates = @('Explicit', 'ImplicitRows', 'ImplicitRowsAndCells')
                }))
    }
    excelpublicparallelread = [pscustomobject]@{
        Area = 'excel'
        Title = '65K workbook sequential and ordered parallel reading'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only sequential and parallel read variants; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelPublicParallelReadBenchmarks.*'
        ComparisonId = 'excel-public-parallel-read'
        Suite = 'OfficeIMO.Excel.PublicParallelRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('Sequential', 'OrderedParallel') -Parameters ([ordered]@{
                    Format = $formats
                }))
    }
    excelpublicparallelprojection = [pscustomobject]@{
        Area = 'excel'
        Title = '65K workbook sequential and ordered parallel projection'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only sequential and parallel projection variants; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelPublicParallelProjectionBenchmarks.*'
        ComparisonId = 'excel-public-parallel-projection'
        Suite = 'OfficeIMO.Excel.PublicParallelProjection'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('Sequential', 'OrderedParallel') -Parameters ([ordered]@{
                    Format = $formats
                    ProjectionRounds = @(4096)
                }))
    }
    excelrawreadstage = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX reader preparation and row consumption stages'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only reader stage profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Frameworks = @('net10.0')
        Filter = '*.ExcelRawReadStageBenchmarks.*'
        ComparisonId = 'excel-raw-read-stage'
        Suite = 'OfficeIMO.Excel.RawReadStage'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('PrepareAndDispose', 'ConsumePreparedRows') -Parameters ([ordered]@{
                    EscapedNames = $booleans
                }))
    }
    excelusedrangediscovery = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX used-range discovery'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only used-range profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelUsedRangeDiscoveryBenchmarks.*'
        ComparisonId = 'excel-used-range-discovery'
        Suite = 'OfficeIMO.Excel.UsedRangeDiscovery'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('DiscoverRange') -Parameters ([ordered]@{
                    RowCount = @(2500, 25000)
                    Sparse = $booleans
                    OmitRowCoordinates = $booleans
                }))
    }
    excelutf8inlinetextread = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX inline text reading'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only inline-text markup profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelUtf8InlineTextReadBenchmarks.*'
        ComparisonId = 'excel-inline-text-read'
        Suite = 'OfficeIMO.Excel.InlineTextRead'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('Read') -Parameters ([ordered]@{
                    RowCount = @(20000)
                    LineEndings = @('LF', 'CRLF')
                    Coordinates = @('Explicit', 'ImplicitRows')
                }))
    }
    excelworksheetcopy = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX worksheet copy between workbooks'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only copy-mode variants; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelWorksheetCopyBenchmarks.*'
        ComparisonId = 'excel-worksheet-copy'
        Suite = 'OfficeIMO.Excel.WorksheetCopy'
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('PackageCopy', 'ValuesCopy') -Parameters ([ordered]@{
                    RowCount = @(100, 2500, 25000)
                }))
    }
    excelworksheetpreparation = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSX worksheet preparation and traversal'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only worksheet preparation profile; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.ExcelWorksheetPreparationBenchmarks.*'
        ComparisonId = 'excel-worksheet-preparation'
        Suite = 'OfficeIMO.Excel.WorksheetPreparation'
        # RowCount and Markup come from OFFICEIMO_TYPED_BENCHMARK_ROWS and OFFICEIMO_WORKSHEET_BENCHMARK_MARKUP;
        # declared runs leave both unset.
        ExpectedCases = @(New-ExcelExpectedCases -Methods @('OpenThroughFirstRow', 'FullScan') -Parameters ([ordered]@{
                    RowCount = @(25000)
                    Dimension = @('Correct', 'Absent', 'StaleNarrow')
                    StoredWorksheet = $booleans
                    EnableWorksheetPrefetch = $booleans
                    Markup = @('Canonical')
                }))
    }
    excelxlsbpipeline = [pscustomobject]@{
        Area = 'excel'
        Title = 'XLSB 65K public reader and worksheet reader'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = $platforms
        PublishRestriction = 'OfficeIMO-only diagnostic of the public XLSB reader against its worksheet reader; not a library comparison.'
        CatalogEligible = $false
        Project = $project
        Filter = '*.XlsbOfficeIMOPipelineBenchmarks.*'
        ComparisonId = 'markpflug-65k-xlsb-officeimo-pipeline'
        Suite = 'OfficeIMO.Excel.Xlsb.MarkPflug65KPipeline'
        ExpectedCases = @('PublicApi', 'DirectWorksheetReader')
    }
}
