# CSV comparisons beyond the shared definitions in Build/LibraryComparisonDefinitions.ps1. Loading this file performs
# no work; Get-LibraryComparisonDefinitions merges the dictionary it returns.
param([ValidateSet('net8.0', 'net10.0')] [string] $Framework = 'net10.0')

[ordered]@{
    csvwrite100k = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV data-reader writing, 100K rows'
        Libraries = @('OfficeIMO', 'Sylvan.Data.Csv')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvDataReaderWriteBenchmarks*(RowCount: 100000,*'
        ComparisonId = 'csv-100k-datareader-write'
        Suite = 'OfficeIMO.CSV.DataReaderWrite100K'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'OfficeIMO_WriteDataReader'
                'OfficeIMO_WriteDataReaderParallel'
                'Sylvan_WriteDataReader'
            )
            foreach ($rowcount in @('100000')) {
                foreach ($shape in @('Mixed', 'Quoted', 'Multiline')) {
                    foreach ($method in $methods) {
                        "$method|RowCount=$rowcount&Shape=$shape"
                    }
                }
            }
        )
    }
    csvtextwrite = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV text and delimiter writing'
        Libraries = @('OfficeIMO', 'CsvHelper')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvTextWriteBenchmarks*'
        ComparisonId = 'csv-text-write'
        Suite = 'OfficeIMO.CSV.TextWrite'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper'
                'OfficeIMO'
            )
            foreach ($textlength in @('64', '4096')) {
                foreach ($quotemode in @('AsNeeded', 'Always')) {
                    foreach ($delimiter in @(',', ';', '\t', '||')) {
                        foreach ($textshape in @('Plain', 'Delimited', 'Notes', 'Json', 'Quotes')) {
                            foreach ($method in $methods) {
                                "$method|TextLength=$textlength&QuoteMode=$quotemode&Delimiter=$delimiter&TextShape=$textshape"
                            }
                        }
                    }
                }
            }
        )
    }
    csvfilewrite = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV UTF-8 file writing'
        Libraries = @('OfficeIMO', 'CsvHelper')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvFileWriteBenchmarks*'
        ComparisonId = 'csv-file-write'
        Suite = 'OfficeIMO.CSV.FileWrite'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper'
                'OfficeIMO'
            )
            foreach ($shape in @('ShortAscii', 'ShortUnicode', 'DenseJson', 'QuoteRuns', 'LongNotes', 'TypedValues', 'MixedJson25K')) {
                foreach ($quotemode in @('AsNeeded', 'Always')) {
                    foreach ($method in $methods) {
                        "$method|Shape=$shape&QuoteMode=$quotemode"
                    }
                }
            }
        )
    }
    csvquotedposition = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV quoted field position reading'
        Libraries = @('OfficeIMO', 'Sep', 'Sylvan.Data.Csv', 'CsvHelper')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvQuotedFieldPositionBenchmarks*'
        ComparisonId = 'csv-quoted-field-position'
        Suite = 'OfficeIMO.CSV.QuotedFieldPosition'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper'
                'OfficeIMO'
                'Sep'
                'Sylvan'
            )
            foreach ($rowcount in @('5000', '50000')) {
                foreach ($delimiter in @(',', '|')) {
                    foreach ($quotedcolumn in @('1', '6', '13')) {
                        foreach ($method in $methods) {
                            "$method|RowCount=$rowcount&Delimiter=$delimiter&QuotedColumn=$quotedcolumn"
                        }
                    }
                }
            }
        )
    }
    csvpackageassets = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV PackageAssets reading'
        Libraries = @('OfficeIMO', 'Sep', 'Sylvan.Data.Csv', 'CsvHelper')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*PackageAssetsCsvBenchmarks*'
        ComparisonId = 'csv-packageassets-read'
        Suite = 'OfficeIMO.CSV.PackageAssets'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper'
                'OfficeIMO'
                'Sep'
                'Sylvan'
            )
            foreach ($shape in @('Plain', 'Quotes')) {
                foreach ($method in $methods) {
                    "$method|Shape=$shape"
                }
            }
        )
    }
    csvpackageassetsspaces = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV PackageAssets trimmed spaces and quotes reading'
        Libraries = @('OfficeIMO', 'Sep', 'CsvHelper')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*PackageAssetsSpacesQuotesCsvBenchmarks*'
        ComparisonId = 'csv-packageassets-spaces-quotes-read'
        Suite = 'OfficeIMO.CSV.PackageAssetsSpacesQuotes'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper'
                'OfficeIMO'
                'Sep'
            )
            foreach ($method in $methods) {
                $method
            }
        )
    }
    csvwritecompare = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV writing from objects, tables and data readers'
        Libraries = @('OfficeIMO', 'CsvHelper', 'Sep', 'Sylvan.Data.Csv', 'ExcelReader.NET')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvWriteComparisonBenchmarks*'
        ComparisonId = 'csv-write-sources'
        Suite = 'OfficeIMO.CSV.WriteSources'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper_DataReaderText'
                'CsvHelper_DataReaderUtf8'
                'CsvHelper_DataTableText'
                'CsvHelper_DataTableUtf8'
                'CsvHelper_ObjectsText'
                'CsvHelper_ObjectsUtf8'
                'OfficeIMO_DataReaderText'
                'OfficeIMO_DataReaderUtf8'
                'OfficeIMO_DataTableText'
                'OfficeIMO_DataTableUtf8'
                'OfficeIMO_ObjectsText'
                'OfficeIMO_ObjectsUtf8'
                'Sep_DataReaderText'
                'Sep_DataReaderUtf8'
                'Sep_DataTableText'
                'Sep_DataTableUtf8'
                'Sylvan_DataReaderText'
                'Sylvan_DataReaderUtf8'
                'Sylvan_DataTableText'
                'Sylvan_DataTableUtf8'
                if ($Framework -eq 'net10.0') { 'ExcelReaderNet_DataReaderUtf8', 'ExcelReaderNet_DataTableUtf8', 'ExcelReaderNet_ObjectsUtf8', 'Sep_ObjectsText', 'Sep_ObjectsUtf8' }
            )
            foreach ($shape in @('QuotedText', 'Numbers', 'Dates', 'Wide')) {
                foreach ($method in $methods) {
                    "$method|Shape=$shape"
                }
            }
        )
    }
    csvparallelreader = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV typed parallel data-reader reading'
        Libraries = @('OfficeIMO', 'Dataplat.Dbatools.Csv')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'Lanes are checked by typed sums, an ordered Id hash and name lengths rather than every decoded value; not published until the check compares values.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvParallelDataReaderBenchmarks*'
        ComparisonId = 'csv-parallel-datareader'
        Suite = 'OfficeIMO.CSV.ParallelDataReader'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        MethodIdentity = @{
            'DataplatSequential' = 'Dataplat.Dbatools.Csv|Sequential'
            'DataplatParallel' = 'Dataplat.Dbatools.Csv|Parallel'
        }
        ExpectedCases = @(
            $methods = @(
                'DataplatParallel'
                'DataplatSequential'
                'OfficeIMOParallel'
                'OfficeIMOSequential'
            )
            foreach ($rowcount in @('100000')) {
                foreach ($maxdegreeofparallelism in @('4')) {
                    foreach ($batchsize in @('4096')) {
                        foreach ($method in $methods) {
                            "$method|RowCount=$rowcount&MaxDegreeOfParallelism=$maxdegreeofparallelism&BatchSize=$batchsize"
                        }
                    }
                }
            }
        )
    }
    csvprojectedparallel = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV projected parallel typed reading'
        Libraries = @('OfficeIMO', 'Sep', 'RecordParser')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvProjectedParallelBenchmarks*'
        ComparisonId = 'csv-projected-parallel'
        Suite = 'OfficeIMO.CSV.ProjectedParallel'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'OfficeIMO'
                'RecordParser'
                'Sep'
            )
            foreach ($degreeofparallelism in @('1', '4', '8', '16')) {
                foreach ($method in $methods) {
                    "$method|DegreeOfParallelism=$degreeofparallelism"
                }
            }
        )
    }
    csvdbatoolsexport = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV million-row data-reader export'
        Libraries = @('OfficeIMO', 'Dataplat.Dbatools.Csv', 'Sylvan.Data.Csv', 'CsvHelper', 'Sep')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvDbatoolsExportBenchmarks.*Reader*'
        ComparisonId = 'csv-dbatools-export'
        Suite = 'OfficeIMO.CSV.DbatoolsExport'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        MethodIdentity = @{
            'Dataplat_WriteFromReader' = 'Dataplat.Dbatools.Csv|DataReader'
        }
        ExpectedCases = @(
            $methods = @(
                'CsvHelper_DataReader'
                'Dataplat_WriteFromReader'
                'OfficeIMO_DataReader'
                'OfficeIMO_DataReaderDefaultFormat'
                'OfficeIMO_PassThroughDataReader'
                'Sylvan_DataReader'
                if ($Framework -eq 'net10.0') { 'Sep_DataReader' }
            )
            foreach ($rowcount in @('1000000')) {
                foreach ($method in $methods) {
                    "$method|RowCount=$rowcount"
                }
            }
        )
    }
    csvdbatoolsimport = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV million-row typed import'
        Libraries = @('OfficeIMO', 'Dataplat.Dbatools.Csv')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'Lanes compare Id, score and date values but names only by length; not published until names are compared.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvDbatoolsExportBenchmarks.*_TypedRead'
        ComparisonId = 'csv-dbatools-typed-import'
        Suite = 'OfficeIMO.CSV.DbatoolsTypedImport'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'Dataplat_TypedRead'
                'OfficeIMO_TypedRead'
            )
            foreach ($rowcount in @('1000000')) {
                foreach ($method in $methods) {
                    "$method|RowCount=$rowcount"
                }
            }
        )
    }
    csvexcelreadergenerated = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV generated 50K read, parse and write'
        Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sylvan.Data.Csv', 'Sep', 'CsvHelper', 'Dataplat.Dbatools.Csv')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvExcelReaderGeneratedBenchmarks*'
        ComparisonId = 'csv-generated-read-parse-write'
        Suite = 'OfficeIMO.CSV.GeneratedReadParseWrite'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper_Parse'
                'CsvHelper_Read'
                'Dataplat_Read'
                'ExcelReaderNet_Parse'
                'ExcelReaderNet_Read'
                'ExcelReaderNet_Write'
                'OfficeIMO_Parse'
                'OfficeIMO_Read'
                'OfficeIMO_Write'
                'Sep_Read'
                'Sylvan_Read'
            )
            foreach ($method in $methods) {
                $method
            }
        )
    }
    csvtrimunescape = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV trim and unescape to strings'
        Libraries = @('OfficeIMO', 'Sep', 'CsvHelper')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvTrimUnescapeBenchmarks*'
        ComparisonId = 'csv-trim-unescape-strings'
        Suite = 'OfficeIMO.CSV.TrimUnescapeStrings'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        MethodIdentity = @{
            'OfficeIMODataReaderStrings' = 'OfficeIMO|Strings'
        }
        ExpectedCases = @(
            $methods = @(
                'CsvHelperStrings'
                'OfficeIMODataReaderStrings'
                'SepStrings'
            )
            foreach ($rowcount in @('50000')) {
                foreach ($method in $methods) {
                    "$method|RowCount=$rowcount"
                }
            }
        )
    }
    csvtrimunescapespans = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV trim and unescape to spans'
        Libraries = @('OfficeIMO', 'Sep')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvTrimUnescapeSpanBenchmarks*'
        ComparisonId = 'csv-trim-unescape-spans'
        Suite = 'OfficeIMO.CSV.TrimUnescapeSpans'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        MethodIdentity = @{
            'OfficeIMOFieldSpans' = 'OfficeIMO|Spans'
        }
        ExpectedCases = @(
            $methods = @(
                'OfficeIMOFieldSpans'
                'SepSpans'
            )
            foreach ($rowcount in @('50000')) {
                foreach ($method in $methods) {
                    "$method|RowCount=$rowcount"
                }
            }
        )
    }
    csvtypedsequential = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV sequential typed materialization'
        Libraries = @('OfficeIMO', 'Sep', 'RecordParser')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvTypedSequentialBenchmarks*'
        ComparisonId = 'csv-typed-sequential'
        Suite = 'OfficeIMO.CSV.TypedSequential'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        MethodIdentity = @{
            'OfficeIMOManual' = 'OfficeIMO|Sequential'
        }
        ExpectedCases = @(
            $methods = @(
                'OfficeIMOManual'
                'RecordParserSequential'
                'SepSequential'
            )
            foreach ($rowcount in @('25000', '100000')) {
                foreach ($method in $methods) {
                    "$method|RowCount=$rowcount"
                }
            }
        )
    }
    csvparallelscaling = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV ordered parallel typed materialization'
        Libraries = @('OfficeIMO', 'Sep', 'RecordParser')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvParallelScalingBenchmarks*'
        ComparisonId = 'csv-typed-parallel-scaling'
        Suite = 'OfficeIMO.CSV.TypedParallelScaling'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'OfficeIMOParallel'
                'RecordParserParallel'
                'SepParallel'
            )
            foreach ($degreeofparallelism in @('1', '4', '8', '16')) {
                foreach ($method in $methods) {
                    "$method|DegreeOfParallelism=$degreeofparallelism"
                }
            }
        )
    }
    csvparallelcrossover = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV 25K parallel typed materialization'
        Libraries = @('OfficeIMO', 'Sep')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'None'
        CatalogEligible = $true
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvParallelCrossoverBenchmarks*'
        ComparisonId = 'csv-typed-parallel-crossover'
        Suite = 'OfficeIMO.CSV.TypedParallelCrossover'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'OfficeIMOParallel'
                'SepParallel'
            )
            foreach ($degreeofparallelism in @('16')) {
                foreach ($method in $methods) {
                    "$method|DegreeOfParallelism=$degreeofparallelism"
                }
            }
        )
    }
    csvautomaticmapping = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV automatic property mapping'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'OfficeIMO-only API variants (automatic mapper against the explicit typed loop); not a library comparison.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvAutomaticMappingBenchmarks*'
        ComparisonId = 'csv-automatic-mapping'
        Suite = 'OfficeIMO.CSV.AutomaticMapping'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'OfficeIMOManual'
                'OfficeIMORowsAs'
            )
            foreach ($rowcount in @('25000', '100000')) {
                foreach ($method in $methods) {
                    "$method|RowCount=$rowcount"
                }
            }
        )
    }
    csvparallelofficetuning = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV parallel batch-size tuning'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'OfficeIMO-only tuning sweep; not a library comparison.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvParallelOfficeTuningBenchmarks*'
        ComparisonId = 'csv-parallel-office-tuning'
        Suite = 'OfficeIMO.CSV.ParallelOfficeTuning'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'OfficeIMORecordParallel'
            )
            foreach ($batchsize in @('1024', '2048', '4096', '8192')) {
                foreach ($degreeofparallelism in @('8', '16')) {
                    foreach ($method in $methods) {
                        "$method|BatchSize=$batchsize&DegreeOfParallelism=$degreeofparallelism"
                    }
                }
            }
        )
    }
    csvparallelcrossovertuning = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV 25K parallel batch-size tuning'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'OfficeIMO-only tuning sweep; not a library comparison.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvParallelCrossoverTuningBenchmarks*'
        ComparisonId = 'csv-parallel-crossover-tuning'
        Suite = 'OfficeIMO.CSV.ParallelCrossoverTuning'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'OfficeIMORecordParallel'
            )
            foreach ($batchsize in @('1024', '2048', '3072', '3584', '4096')) {
                foreach ($degreeofparallelism in @('16')) {
                    foreach ($method in $methods) {
                        "$method|BatchSize=$batchsize&DegreeOfParallelism=$degreeofparallelism"
                    }
                }
            }
        )
    }
    csvtsvdialect = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV and TSV dialect reading'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'OfficeIMO-only dialect variants; not a library comparison.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvTsvDialectBenchmarks*'
        ComparisonId = 'csv-tsv-dialect'
        Suite = 'OfficeIMO.CSV.TsvDialect'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'Csv_FieldSpans'
                'Csv_PathReader'
                'Csv_TextReader'
                'Tsv_FieldSpans'
                'Tsv_PathReader'
                'Tsv_TextReader'
            )
            foreach ($rowcount in @('100000')) {
                foreach ($method in $methods) {
                    "$method|RowCount=$rowcount"
                }
            }
        )
    }
    csvdocumentsave = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV document save and compression'
        Libraries = @('OfficeIMO')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'OfficeIMO-only save and compression paths; not a library comparison.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvDocumentSaveBenchmarks*'
        ComparisonId = 'csv-document-save'
        Suite = 'OfficeIMO.CSV.DocumentSave'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'RowWriterFile'
                'Save'
                'SaveAsync'
                'SaveFile'
                'SaveFileAsync'
                'ToBytes'
                'ToText'
                'WriteDataReader'
                'WriteDataReaderParallel'
            )
            foreach ($rowcount in @('1000', '25000')) {
                foreach ($shape in @('Plain', 'Quoted', 'MixedJson')) {
                    foreach ($compression in @('None', 'GZip', 'Deflate', 'Brotli', 'ZLib')) {
                        foreach ($method in $methods) {
                            "$method|RowCount=$rowcount&Shape=$shape&Compression=$compression"
                        }
                    }
                }
            }
        )
    }
    csvdbatoolsparity = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV dbatools reader parity'
        Libraries = @('OfficeIMO', 'Dataplat.Dbatools.Csv', 'LumenWorks', 'Sep', 'Sylvan.Data.Csv', 'CsvHelper')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'Lanes check only row counts and total field lengths, not decoded values, and the groups use different baselines; not published until values are compared.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvDbatoolsLibraryParityBenchmarks*'
        ComparisonId = 'csv-dbatools-parity'
        Suite = 'OfficeIMO.CSV.DbatoolsParity'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper_Large'
                'CsvHelper_Medium'
                'CsvHelper_QuickTest_AllColumns'
                'CsvHelper_QuickTest_SingleColumn'
                'Dataplat_AllValues'
                'Dataplat_Large'
                'Dataplat_Large_Modern'
                'Dataplat_Medium'
                'Dataplat_Medium_Modern'
                'Dataplat_QuickTest_AllColumns'
                'Dataplat_QuickTest_SingleColumn'
                'Dataplat_Quoted'
                'Dataplat_Small'
                'Dataplat_Wide'
                'LumenWorks_AllValues'
                'LumenWorks_Large'
                'LumenWorks_Medium'
                'LumenWorks_QuickTest_AllColumns'
                'LumenWorks_QuickTest_SingleColumn'
                'LumenWorks_Quoted'
                'LumenWorks_Small'
                'LumenWorks_Wide'
                'OfficeIMO_AllValues'
                'OfficeIMO_DataReader_AllValues'
                'OfficeIMO_DataReader_QuickTest_GetValues'
                'OfficeIMO_Large'
                'OfficeIMO_Large_Modern'
                'OfficeIMO_Medium'
                'OfficeIMO_Medium_Modern'
                'OfficeIMO_QuickTest_AllColumns'
                'OfficeIMO_QuickTest_SingleColumn'
                'OfficeIMO_Quoted'
                'OfficeIMO_Small'
                'OfficeIMO_Wide'
                'Sep_Large'
                'Sep_Medium'
                'Sep_QuickTest_AllColumns'
                'Sep_QuickTest_SingleColumn'
                'Sylvan_Large'
                'Sylvan_Medium'
                'Sylvan_QuickTest_AllColumns'
                'Sylvan_QuickTest_SingleColumn'
            )
            foreach ($method in $methods) {
                $method
            }
        )
    }
    csvapimatrix = [pscustomobject]@{
        Area = 'csv'
        Title = 'CSV read and write API matrix'
        Libraries = @('OfficeIMO', 'CsvHelper', 'ExcelReader.NET', 'Sylvan.Data.Csv', 'Dataplat.Dbatools.Csv', 'Sep')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'Exploratory API matrix: most cases exist for one library only and validation is not recorded per lane; the shared operations are published by the focused CSV comparisons.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*.CsvBenchmarks.*'
        ComparisonId = 'csv-api-matrix'
        Suite = 'OfficeIMO.CSV.ApiMatrix'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper_ReadFields'
                'CsvHelper_ReadTypedRecords'
                'CsvHelper_WriteProjectedRows'
                'CsvHelper_WriteTextRows'
                'CsvHelper_WriteTypedRecords'
                'Dataplat_ReadDataTableLoad'
                'Dataplat_ReadFields'
                'Dataplat_WriteFromReader'
                'Dataplat_WriteProjectedRows'
                'Dataplat_WriteTextRows'
                'ExcelReaderNet_ReadFields'
                'ExcelReaderNet_ReadTypedRecords'
                'ExcelReaderNet_WriteProjectedRows'
                'ExcelReaderNet_WriteTextRows'
                'ExcelReaderNet_WriteTypedRecords'
                'OfficeIMO_ReadDataReaderGetStrings'
                'OfficeIMO_ReadDataReaderInferredSchema'
                'OfficeIMO_ReadDataReaderStrings'
                'OfficeIMO_ReadDataTableInferredSchema'
                'OfficeIMO_ReadDataTableLoadDataReaderInferredSchema'
                'OfficeIMO_ReadDataTableLoadDataReaderStrings'
                'OfficeIMO_ReadDataTableStrings'
                'OfficeIMO_ReadFieldSpanVisitorSkipHeader'
                'OfficeIMO_ReadFieldSpansMaterialized'
                'OfficeIMO_ReadInMemoryRows'
                'OfficeIMO_ReadRowFieldSpansMaterialized'
                'OfficeIMO_ReadRowsCallback'
                'OfficeIMO_ReadRowsReusableCallback'
                'OfficeIMO_ReadStreamGetStrings'
                'OfficeIMO_ReadStreamingDataTableInferredSchemaDefaultSample'
                'OfficeIMO_ReadStreamingDataTableStrings'
                'OfficeIMO_ReadStreamingRows'
                'OfficeIMO_ReadTextFieldSpanVisitorSkipHeader'
                'OfficeIMO_ReadTextRowFieldSpansMaterialized'
                'OfficeIMO_ReadTypedRowsForwardOnly'
                'OfficeIMO_ReadTypedRowsMaterialized'
                'OfficeIMO_WriteDataReader'
                'OfficeIMO_WriteIncrementalProjectedRows'
                'OfficeIMO_WriteObjects'
                'OfficeIMO_WriteProjectedRows'
                'OfficeIMO_WriteProjectedRowsAlwaysQuoted'
                'OfficeIMO_WriteProjectedRowsUtf8'
                'OfficeIMO_WriteTextRows'
                'OfficeIMO_WriteTypedRecordsUtf8'
                'OfficeIMO_WriteValidatedTextRows'
                'Sep_ReadFieldSpans'
                'Sep_ReadFields'
                'Sep_WriteTextRows'
                'Sylvan_ReadDataTableLoad'
                'Sylvan_ReadFieldSpans'
                'Sylvan_ReadFields'
                'Sylvan_WriteProjectedRows'
                'Sylvan_WriteTextRows'
            )
            foreach ($rowcount in @('1000', '10000', '25000')) {
                foreach ($shape in @('Mixed', 'Quoted', 'Multiline')) {
                    foreach ($method in $methods) {
                        "$method|RowCount=$rowcount&Shape=$shape"
                    }
                }
            }
        )
    }
    csvwideapimatrix = [pscustomobject]@{
        Area = 'csv'
        Title = 'Wide CSV read and write API matrix'
        Libraries = @('OfficeIMO', 'CsvHelper', 'ExcelReader.NET', 'Sylvan.Data.Csv', 'Dataplat.Dbatools.Csv', 'Sep')
        ExpectedPlatforms = @('windows', 'linux', 'macos')
        PublishRestriction = 'Exploratory API matrix: most cases exist for one library only and validation is not recorded per lane; the shared operations are published by the focused CSV comparisons.'
        CatalogEligible = $false
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*CsvWideBenchmarks*'
        ComparisonId = 'csv-wide-api-matrix'
        Suite = 'OfficeIMO.CSV.WideApiMatrix'
        FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
        ExpectedCases = @(
            $methods = @(
                'CsvHelper_ReadFields'
                'CsvHelper_WriteProjectedRows'
                'CsvHelper_WriteTextRows'
                'Dataplat_ReadDataTableLoad'
                'Dataplat_ReadFields'
                'Dataplat_WriteFromReader'
                'Dataplat_WriteProjectedRows'
                'Dataplat_WriteTextRows'
                'ExcelReaderNet_ReadUtf8Fields'
                'ExcelReaderNet_WriteProjectedRowsUtf8Stream'
                'OfficeIMO_ReadDataReaderExplicitSchema'
                'OfficeIMO_ReadDataReaderGetStrings'
                'OfficeIMO_ReadDataReaderInferredSchema'
                'OfficeIMO_ReadDataReaderStrings'
                'OfficeIMO_ReadDataTableInferredSchema'
                'OfficeIMO_ReadDataTableLoadDataReaderInferredSchema'
                'OfficeIMO_ReadDataTableLoadDataReaderStrings'
                'OfficeIMO_ReadDataTableStrings'
                'OfficeIMO_ReadFieldSpanVisitorSkipHeader'
                'OfficeIMO_ReadFieldSpansMaterializedSkipHeader'
                'OfficeIMO_ReadFieldSpansSkipHeader'
                'OfficeIMO_ReadFileDataReaderExplicitSchema'
                'OfficeIMO_ReadRecordsReusableSkipHeader'
                'OfficeIMO_ReadRowFieldSpansMaterialized'
                'OfficeIMO_ReadRowsReusableCallback'
                'OfficeIMO_ReadTextFieldSpanVisitorSkipHeader'
                'OfficeIMO_ReadTextRowFieldSpansMaterialized'
                'OfficeIMO_ReadUtf8Fields'
                'OfficeIMO_WriteDataReader'
                'OfficeIMO_WriteIncrementalProjectedRows'
                'OfficeIMO_WriteProjectedRows'
                'OfficeIMO_WriteProjectedRowsUtf8Stream'
                'OfficeIMO_WriteTextRows'
                'OfficeIMO_WriteValidatedTextRows'
                'Sep_ReadFieldSpans'
                'Sep_ReadFields'
                'Sep_WriteTextRows'
                'Sylvan_ReadDataTableLoad'
                'Sylvan_ReadFieldSpans'
                'Sylvan_ReadFields'
                'Sylvan_WriteProjectedRows'
                'Sylvan_WriteTextRows'
            )
            foreach ($rowcount in @('1000', '10000', '25000')) {
                foreach ($method in $methods) {
                    "$method|RowCount=$rowcount"
                }
            }
        )
    }
}
