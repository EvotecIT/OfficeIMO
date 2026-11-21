# Shared workload declarations for collection and site generation. Loading this file performs no work.
# Further suites are declared per area in Build/LibraryComparisonDefinitions/<area>.ps1; each file takes -Framework
# and returns an ordered dictionary of workloads in the same shape.
# PublishRestriction records why a comparison's shape is unpublishable. Licence restrictions come from the
# libraries in Build/benchmark-libraries.json; the runner derives publishability from both.
function Get-LibraryComparisonDefinitions {
    <#
    .SYNOPSIS
    Returns every library-comparison declaration without executing a runner.
    .PARAMETER Framework
    Target framework. It selects the cases compiled for that framework; comparison identities never include it.
    .EXAMPLE
    Get-LibraryComparisonDefinitions -Framework net10.0
    .OUTPUTS
    An ordered dictionary keyed by workload, including restricted and uncollected lanes.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('net8.0', 'net10.0')]
        [string] $Framework = 'net10.0'
    )

    # The Recursive AI One libraries (OnePDF, OneDocument, OneRender, OneImage, OneMail, OneArchive) ship net10.0
    # only, so their cases and library names are declared only for that target.
    $one = $Framework -eq 'net10.0'
    # Full JSON exports import Scenario from the CLR Method, not Benchmark.Description.
    $definitions = [ordered]@{
        csvlargefirstrow = [pscustomobject]@{
            Area = 'csv'
            Title = 'CSV million-row first-row access'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO-only snapshot/incremental profile; excluded from public library comparisons.'
            Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
            Filter = '*CsvLargeFirstRowBenchmarks*'
            ComparisonId = 'csv-million-first-row'
            Suite = 'OfficeIMO.CSV.LargeFirstRow'
            CatalogEligible = $false
            FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
            ExpectedCases = @(
                foreach ($shape in @('Plain', 'Multiline')) {
                    foreach ($scenario in @('OfficeIMO_Snapshot', 'OfficeIMO_Incremental')) {
                        "$scenario|Shape=$shape"
                    }
                }
            )
        }
        csvasyncread = [pscustomobject]@{
            Area = 'csv'
            Title = 'CSV asynchronous reading'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO-only snapshot/incremental profile; excluded from public library comparisons.'
            Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
            Filter = '*CsvAsyncReadBenchmarks*'
            ComparisonId = 'csv-async-read'
            Suite = 'OfficeIMO.CSV.AsyncRead'
            CatalogEligible = $false
            FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
            ExpectedCases = @(
                foreach ($operation in @('FirstRow', 'AllRows')) {
                    foreach ($rowCount in @(1000, 100000)) {
                        foreach ($shape in @('Plain', 'Multiline')) {
                            foreach ($scenario in @('OfficeIMO_Snapshot', 'OfficeIMO_Incremental')) {
                                "$scenario|Operation=$operation&RowCount=$rowCount&Shape=$shape"
                            }
                        }
                    }
                }
            )
        }
        csvwideasyncread = [pscustomobject]@{
            Area = 'csv'
            Title = 'Wide CSV asynchronous reading'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO-only snapshot/incremental profile; excluded from public library comparisons.'
            Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
            Filter = '*CsvWideAsyncReadBenchmarks*'
            ComparisonId = 'csv-wide-async-read'
            Suite = 'OfficeIMO.CSV.WideAsyncRead'
            CatalogEligible = $false
            FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
            ExpectedCases = @(
                foreach ($operation in @('FirstRow', 'AllRows')) {
                    foreach ($rowCount in @(5000, 25000)) {
                        foreach ($scenario in @('OfficeIMO_Snapshot', 'OfficeIMO_Incremental')) {
                            "$scenario|Operation=$operation&RowCount=$rowCount"
                        }
                    }
                }
            )
        }
        pdfnative = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF native operations'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfNativeOperationsBenchmarks*'
            ComparisonId = 'pdf-native-100-pages'
            Suite = 'OfficeIMO.Pdf.NativeOperations'
            ExpectedCases = @(
                foreach ($producer in @('IText', 'MigraDoc')) {
                    foreach ($operation in @('ReadText', 'ReadStructured', 'Select', 'Merge', 'Split', 'SplitSelections')) {
                        "Execute|Operation=$operation&Producer=$producer"
                    }
                }
            )
        }
        csv = [pscustomobject]@{
            Area = 'csv'
            Title = 'CSV 65K typed reading'
            Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sep', 'Sylvan.Data.Csv', 'CsvHelper', 'Dataplat.Dbatools.Csv', 'LumenWorks')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
            Filter = '*MarkPflug65KCsvBenchmarks*'
            ComparisonId = 'markpflug-65k-csv-decoded'
            Suite = 'OfficeIMO.CSV.MarkPflug65K'
            ExpectedCases = @('OfficeIMO', 'ExcelReaderNet', 'Sep', 'Sylvan', 'CsvHelper', 'DataplatDbatools', 'LumenWorks')
        }
        csvwrite = [pscustomobject]@{
            Area = 'csv'
            Title = 'CSV data-reader writing'
            Libraries = @('OfficeIMO', 'Sylvan.Data.Csv')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
            Filter = '*CsvDataReaderWriteBenchmarks*(RowCount: 25000,*'
            ComparisonId = 'csv-25k-datareader-write'
            Suite = 'OfficeIMO.CSV.DataReaderWrite25K'
            ExpectedCases = @(
                foreach ($shape in @('Mixed', 'Quoted', 'Multiline')) {
                    foreach ($scenario in @(
                        'OfficeIMO_WriteDataReader'
                        'OfficeIMO_WriteDataReaderParallel'
                        'Sylvan_WriteDataReader')) {
                        "$scenario|RowCount=25000&Shape=$shape"
                    }
                }
            )
        }
        xls = [pscustomobject]@{
            Area = 'excel'
            Title = 'XLS 65K typed reading'
            Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sylvan.Data.Excel', 'ExcelDataReader'; if ($one) { 'OneDocument.Spreadsheets' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
            Filter = '*MarkPflug65KXlsBenchmarks*'
            ComparisonId = 'markpflug-65k-xls-typed'
            Suite = 'OfficeIMO.Excel.Xls.MarkPflug65K'
            ExpectedCases = @('OfficeIMO', 'ExcelReaderNet', 'Sylvan', 'ExcelDataReader'; if ($one) { 'OneDocument' })
        }
        xlsx = [pscustomobject]@{
            Area = 'excel'
            Title = 'XLSX 65K typed reading'
            # EPPlus is restricted by the library policy: published and default runs drop its case, and a run that
            # enables it requires the case to succeed.
            Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sylvan.Data.Excel', 'ExcelDataReader', 'ClosedXML', 'EPPlus', 'MiniExcel', 'Excel-PRIME', 'HypeLab.IO.Excel', 'XLibur'; if ($one) { 'OneDocument.Spreadsheets' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
            Filter = '*MarkPflug65KXlsxBenchmarks*'
            ComparisonId = 'markpflug-65k-xlsx-typed'
            Suite = 'OfficeIMO.Excel.Xlsx.MarkPflug65K'
            ExpectedCases = @(
                'OfficeIMO'
                'OfficeIMO_Prefetch'
                'ExcelReaderNet'
                'ExcelReaderNet_Prefetch'
                'Sylvan'
                'ExcelDataReader'
                'ClosedXML'
                'EPPlus'
                'MiniExcel'
                'OfficeIMO_DefaultNumbers'
                'OfficeIMO_DefaultNumbers_Prefetch'
                'OfficeIMO_Loaded'
                'OfficeIMO_SinglePass'
                'OfficeIMO_Memory'
                # These factories/options are compiled only for the modern ExcelReader.NET API; OneDocument ships net10.0 only.
                if ($Framework -eq 'net10.0') {
                    'ExcelReaderNet_Memory'
                    'ExcelReaderNet_InternStrings'
                    'OneDocument'
                }
                'ExcelPrime'
                'ExcelPrime_TypedStyles'
                'HypeLab'
                'XLibur'
            )
        }
        xlsxwrite = [pscustomobject]@{
            Area = 'excel'
            Title = 'XLSX data-reader writing'
            Libraries = @('OfficeIMO', 'SpreadCheetah', 'Sylvan.Data.Excel', 'LargeXlsx'; if ($one) { 'OneDocument.Spreadsheets' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
            Filter = '*ExcelDataReaderWriteBenchmarks*'
            ComparisonId = 'xlsx-25k-datareader-write'
            Suite = 'OfficeIMO.Excel.DataReaderWrite25K'
            ExpectedCases = @(
                'OfficeIMO|RowCount=25000'
                'SpreadCheetah|RowCount=25000'
                'Sylvan|RowCount=25000'
                'LargeXlsx|RowCount=25000'
                if ($one) { 'OneDocument|RowCount=25000' }
            )
        }
        xlsb = [pscustomobject]@{
            Area = 'excel'
            Title = 'XLSB 65K typed reading'
            Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sylvan.Data.Excel', 'ExcelDataReader', 'Excel-PRIME'; if ($one) { 'OneDocument.Spreadsheets' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
            Filter = '*MarkPflug65KXlsbBenchmarks*'
            ComparisonId = 'markpflug-65k-xlsb-typed'
            Suite = 'OfficeIMO.Excel.Xlsb.MarkPflug65K'
            ExpectedCases = @(
                'OfficeIMO'
                'ExcelReaderNet'
                'Sylvan'
                'ExcelDataReader'
                'OfficeIMO_SinglePass'
                'ExcelPrime'
                'ExcelPrime_TypedStyles'
                if ($Framework -eq 'net10.0') {
                    'ExcelReaderNet_Prefetch'
                    'OneDocument'
                }
            )
        }
        markdownparse = [pscustomobject]@{
            Area = 'markdown'
            Title = 'CommonMark semantic parsing'
            Libraries = @('OfficeIMO', 'Markdig')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Markdown.Benchmarks\OfficeIMO.Markdown.Benchmarks.csproj'
            Filter = '*MarkdownParseBenchmarks.*_ParseSemantic*_CommonMark'
            ComparisonId = 'markdown-commonmark-parse'
            Suite = 'OfficeIMO.Markdown.CommonMarkParse'
            ExpectedCases = @(
                foreach ($corpus in @(
                    'PortableReadme',
                    'Transcript',
                    'TechnicalDoc',
                    'RichAst',
                    'LongNestedList',
                    'LargeTable',
                    'NormalizationStress')) {
                    foreach ($engine in @('OfficeIMO_ParseSemantic_CommonMark', 'Markdig_ParseSemanticComparison_CommonMark')) {
                        "$engine|CorpusName=$corpus"
                    }
                }
            )
        }
        markdownhtml = [pscustomobject]@{
            Area = 'markdown'
            Title = 'CommonMark to HTML'
            Libraries = @('OfficeIMO', 'Markdig')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Markdown.Benchmarks\OfficeIMO.Markdown.Benchmarks.csproj'
            Filter = '*MarkdownHtmlBenchmarks.*_ToHtml_CommonMark'
            ComparisonId = 'markdown-commonmark-html'
            Suite = 'OfficeIMO.Markdown.CommonMarkHtml'
            ExpectedCases = @(
                foreach ($corpus in @(
                    'PortableReadme',
                    'Transcript',
                    'TechnicalDoc',
                    'RichAst',
                    'LongNestedList',
                    'LargeTable',
                    'NormalizationStress')) {
                    foreach ($engine in @('OfficeIMO_ToHtml_CommonMark', 'Markdig_ToHtml_CommonMark')) {
                        "$engine|CorpusName=$corpus"
                    }
                }
            )
        }
        htmltomarkdown = [pscustomobject]@{
            Area = 'html'
            Title = 'HTML to Markdown'
            Libraries = @('OfficeIMO', 'ReverseMarkdown')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Markdown.Benchmarks\OfficeIMO.Markdown.Benchmarks.csproj'
            Filter = '*HtmlToMarkdownBenchmarks*'
            ComparisonId = 'html-to-markdown'
            Suite = 'OfficeIMO.Markdown.HtmlToMarkdown'
            ExpectedCases = @(
                foreach ($corpus in @('Article', 'LargeArticle', 'Table')) {
                    foreach ($engine in @('OfficeIMO_Default_Profile', 'ReverseMarkdown_Default_Profile')) {
                        "$engine|CorpusName=$corpus"
                    }
                }
            )
        }
        rtfhtml = [pscustomobject]@{
            Area = 'rtf'
            Title = 'RTF to HTML'
            Libraries = @('OfficeIMO', 'RtfPipe')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Rtf.Benchmarks.Comparisons\OfficeIMO.Rtf.Benchmarks.Comparisons.csproj'
            Filter = '*RtfToHtmlComparisonBenchmarks*'
            ComparisonId = 'rtf-to-html'
            Suite = 'OfficeIMO.Rtf.HtmlComparison'
            ValidatedSizeEvidence = $true
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Medium', 'Large', 'Producer')) {
                    foreach ($engine in @('OfficeIMO', 'RtfPipe')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        odscreate = [pscustomobject]@{
            Area = 'opendocument'
            Title = 'ODS creation'
            Libraries = @('OfficeIMO', 'OpenStandardLibrary')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.OpenDocument.Benchmarks.Comparisons\OfficeIMO.OpenDocument.Benchmarks.Comparisons.csproj'
            Filter = '*OdsCreateComparisonBenchmarks*'
            ComparisonId = 'opendocument-ods-create'
            Suite = 'OfficeIMO.OpenDocument.OdsCreate'
            ValidatedSizeEvidence = $true
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal')) {
                    foreach ($engine in @('OfficeIMO', 'OpenStandardLibrary')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        odsread = [pscustomobject]@{
            Area = 'opendocument'
            Title = 'ODS reading'
            Libraries = @('OfficeIMO', 'OpenStandardLibrary')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.OpenDocument.Benchmarks.Comparisons\OfficeIMO.OpenDocument.Benchmarks.Comparisons.csproj'
            Filter = '*OdsReadComparisonBenchmarks*'
            ComparisonId = 'opendocument-ods-read'
            Suite = 'OfficeIMO.OpenDocument.OdsRead'
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal')) {
                    foreach ($engine in @('OfficeIMO', 'OpenStandardLibrary')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        emailmimeread = [pscustomobject]@{
            Area = 'email'
            Title = 'MIME email reading'
            Libraries = @('OfficeIMO', 'MimeKit'; if ($one) { 'OneMail' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Email.Benchmarks.Comparisons\OfficeIMO.Email.Benchmarks.Comparisons.csproj'
            Filter = '*EmailMimeReadBenchmarks*'
            ComparisonId = 'email-mime-read'
            Suite = 'OfficeIMO.Email.MimeRead'
            ValidatedSizeEvidence = $true
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal')) {
                    foreach ($engine in @('OfficeIMO', 'MimeKit'; if ($one) { 'OneMail' })) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        emailmimewrite = [pscustomobject]@{
            Area = 'email'
            Title = 'MIME email writing'
            Libraries = @('OfficeIMO', 'MimeKit'; if ($one) { 'OneMail' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Email.Benchmarks.Comparisons\OfficeIMO.Email.Benchmarks.Comparisons.csproj'
            Filter = '*EmailMimeWriteBenchmarks*'
            ComparisonId = 'email-mime-write'
            Suite = 'OfficeIMO.Email.MimeWrite'
            ValidatedSizeEvidence = $true
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal')) {
                    foreach ($engine in @('OfficeIMO', 'MimeKit'; if ($one) { 'OneMail' })) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        epubread = [pscustomobject]@{
            Area = 'epub'
            Title = 'EPUB opening and reading'
            Libraries = @('OfficeIMO', 'VersOne.Epub')
            # VersOne.Epub's timed workflow uses HtmlAgilityPack; its package is recorded with the run, not measured on its own.
            Dependencies = @('HtmlAgilityPack')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Epub.Benchmarks.Comparisons\OfficeIMO.Epub.Benchmarks.Comparisons.csproj'
            Filter = '*EpubReadComparisonBenchmarks*'
            ComparisonId = 'epub-open-read'
            Suite = 'OfficeIMO.Epub.OpenRead'
            ValidatedSizeEvidence = $true
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal')) {
                    foreach ($engine in @('OfficeIMO', 'VersOneEpub')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        ziptraverse = [pscustomobject]@{
            Area = 'zip'
            Title = 'ZIP guarded metadata traversal'
            Libraries = @('OfficeIMO', 'System.IO.Compression'; if ($one) { 'OneArchive' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO enforces path and expansion safety; the platform lane projects raw metadata without equivalent safety policy. Not a public like-for-like safety comparison.'
            Project = 'OfficeIMO.Archives.Benchmarks\OfficeIMO.Archives.Benchmarks.csproj'
            Filter = '*ZipTraversalComparisonBenchmarks*'
            ComparisonId = 'zip-safe-traversal'
            Suite = 'OfficeIMO.Archives.SafeZipTraversal'
            CatalogEligible = $false
            CollectNonPublishedEvidence = $true
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal', 'Large')) {
                    foreach ($engine in @('OfficeIMO', 'Platform'; if ($one) { 'OneArchive' })) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        yamlparse = [pscustomobject]@{
            Area = 'yaml'
            Title = 'YAML parse to events'
            Libraries = @('OfficeIMO', 'YamlDotNet', 'SharpYaml', 'VYaml')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Yaml.Benchmarks\OfficeIMO.Yaml.Benchmarks.csproj'
            Filter = '*YamlParseEventsBenchmarks*'
            ComparisonId = 'yaml-parse-events'
            Suite = 'OfficeIMO.Yaml.ParseEvents'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'KubernetesStream', 'DeepNesting', 'BlockScalars', 'FlowHeavy', 'AnchorHeavy', 'ManySmallFiles')) {
                    foreach ($engine in @('YamlDotNet', 'SharpYaml', 'OfficeIMOYaml', 'VYaml')) {
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        # VYaml resolves the plain scalar 'infinity' in one corpus file as a float, so the benchmark's lane filter leaves its ManySmallFiles load out.
        yamlload = [pscustomobject]@{
            Area = 'yaml'
            Title = 'YAML load to model'
            Libraries = @('OfficeIMO', 'YamlDotNet', 'SharpYaml', 'VYaml', 'OneYaml')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Yaml.Benchmarks\OfficeIMO.Yaml.Benchmarks.csproj'
            Filter = '*YamlLoadTreeBenchmarks*'
            ComparisonId = 'yaml-load'
            Suite = 'OfficeIMO.Yaml.Load'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'KubernetesStream', 'DeepNesting', 'BlockScalars', 'FlowHeavy', 'AnchorHeavy', 'ManySmallFiles')) {
                    foreach ($engine in @('OneYaml', 'VYaml', 'YamlDotNet', 'SharpYaml', 'OfficeIMOYaml', 'OfficeIMOYamlWithValues')) {
                        if ($engine -eq 'VYaml' -and $workload -eq 'ManySmallFiles') { continue }
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        yamlloadsave = [pscustomobject]@{
            Area = 'yaml'
            Title = 'YAML unchanged load and save'
            Libraries = @('OfficeIMO', 'YamlDotNet', 'SharpYaml', 'OneYaml')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Yaml.Benchmarks\OfficeIMO.Yaml.Benchmarks.csproj'
            Filter = '*YamlLoadAndSaveBenchmarks*'
            ComparisonId = 'yaml-load-save'
            Suite = 'OfficeIMO.Yaml.LoadAndSave'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'KubernetesStream', 'DeepNesting', 'BlockScalars', 'FlowHeavy', 'AnchorHeavy', 'ManySmallFiles')) {
                    foreach ($engine in @('YamlDotNet', 'SharpYaml', 'OneYaml', 'OfficeIMOYaml')) {
                        # SharpYaml writes %TAG directives before each later document without the "..." end marker
                        # YAML requires there, so its multi-document output does not parse in another reader.
                        if ($engine -eq 'SharpYaml' -and $workload -in @('KubernetesStream', 'DeepNesting', 'ManySmallFiles')) { continue }
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        yamledit = [pscustomobject]@{
            Area = 'yaml'
            Title = 'YAML single-value edit and save'
            Libraries = @('OfficeIMO', 'YamlDotNet', 'SharpYaml', 'OneYaml')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Yaml.Benchmarks\OfficeIMO.Yaml.Benchmarks.csproj'
            Filter = '*YamlEditAndSaveBenchmarks*'
            ComparisonId = 'yaml-edit-save'
            Suite = 'OfficeIMO.Yaml.EditAndSave'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'KubernetesStream', 'DeepNesting', 'BlockScalars', 'FlowHeavy', 'AnchorHeavy', 'ManySmallFiles')) {
                    foreach ($engine in @('YamlDotNet', 'SharpYaml', 'OneYaml', 'OfficeIMOYaml')) {
                        # SharpYaml writes %TAG directives before each later document without the "..." end marker
                        # YAML requires there, so its multi-document output does not parse in another reader.
                        if ($engine -eq 'SharpYaml' -and $workload -in @('KubernetesStream', 'DeepNesting', 'ManySmallFiles')) { continue }
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        yamlemit = [pscustomobject]@{
            Area = 'yaml'
            Title = 'YAML emit from events'
            Libraries = @('OfficeIMO', 'YamlDotNet', 'SharpYaml')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Yaml.Benchmarks\OfficeIMO.Yaml.Benchmarks.csproj'
            Filter = '*YamlEmitBenchmarks*'
            ComparisonId = 'yaml-emit'
            Suite = 'OfficeIMO.Yaml.Emit'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'KubernetesStream', 'DeepNesting', 'BlockScalars', 'FlowHeavy', 'AnchorHeavy', 'ManySmallFiles')) {
                    foreach ($engine in @('YamlDotNet', 'SharpYaml', 'OfficeIMOYaml')) {
                        # SharpYaml writes %TAG directives before each later document without the "..." end marker
                        # YAML requires there, so its multi-document output does not parse in another reader.
                        if ($engine -eq 'SharpYaml' -and $workload -in @('KubernetesStream', 'DeepNesting', 'ManySmallFiles')) { continue }
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        yamlserialize = [pscustomobject]@{
            Area = 'yaml'
            Title = 'YAML object serialization'
            Libraries = @('OfficeIMO', 'YamlDotNet', 'SharpYaml', 'VYaml', 'OneYaml')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Yaml.Benchmarks\OfficeIMO.Yaml.Benchmarks.csproj'
            Filter = '*YamlSerializeBenchmarks*'
            ComparisonId = 'yaml-serialize'
            Suite = 'OfficeIMO.Yaml.Serialize'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'LargeRecords', 'KubernetesNested')) {
                    foreach ($engine in @('YamlDotNet', 'SharpYamlGenerated', 'VYamlGenerated', 'OneYamlGenerated', 'OneYamlReflection', 'OneYamlGeneratedUtf8', 'OfficeIMOGenerated', 'OfficeIMOReflection', 'OfficeIMOGeneratedUtf8', 'OfficeIMOGeneratedUtf8Bytes')) {
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        yamldeserialize = [pscustomobject]@{
            Area = 'yaml'
            Title = 'YAML object deserialization'
            Libraries = @('OfficeIMO', 'YamlDotNet', 'SharpYaml', 'VYaml', 'OneYaml')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Yaml.Benchmarks\OfficeIMO.Yaml.Benchmarks.csproj'
            Filter = '*YamlDeserializeBenchmarks*'
            ComparisonId = 'yaml-deserialize'
            Suite = 'OfficeIMO.Yaml.Deserialize'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'LargeRecords', 'KubernetesNested')) {
                    foreach ($engine in @('YamlDotNet', 'SharpYamlGenerated', 'VYamlGenerated', 'OneYamlGenerated', 'OneYamlReflection', 'OneYamlGeneratedUtf8', 'OfficeIMOGenerated', 'OfficeIMOReflection')) {
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        tomlparse = [pscustomobject]@{
            Area = 'toml'
            Title = 'TOML parse without a model'
            Libraries = @('OfficeIMO', 'Tomlyn')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Toml.Benchmarks\OfficeIMO.Toml.Benchmarks.csproj'
            Filter = '*TomlParseBenchmarks*'
            ComparisonId = 'toml-parse'
            Suite = 'OfficeIMO.Toml.Parse'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'CargoWorkspace', 'LockFile', 'Corpus')) {
                    foreach ($engine in @('Tomlyn', 'OfficeIMOToml')) {
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        tomlload = [pscustomobject]@{
            Area = 'toml'
            Title = 'TOML load to model'
            Libraries = @('OfficeIMO', 'Tomlyn', 'Tomlet', 'CsToml')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Toml.Benchmarks\OfficeIMO.Toml.Benchmarks.csproj'
            Filter = '*TomlLoadBenchmarks*'
            ComparisonId = 'toml-load'
            Suite = 'OfficeIMO.Toml.Load'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'CargoWorkspace', 'LockFile', 'Corpus')) {
                    foreach ($engine in @('Tomlyn', 'Tomlet', 'CsToml', 'OfficeIMOToml', 'OfficeIMOTomlWithValues')) {
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        tomlloadsave = [pscustomobject]@{
            Area = 'toml'
            Title = 'TOML unchanged load and save'
            Libraries = @('OfficeIMO', 'Tomlyn')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Toml.Benchmarks\OfficeIMO.Toml.Benchmarks.csproj'
            Filter = '*TomlLoadAndSaveBenchmarks*'
            ComparisonId = 'toml-load-save'
            Suite = 'OfficeIMO.Toml.LoadAndSave'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'CargoWorkspace', 'LockFile', 'Corpus')) {
                    foreach ($engine in @('TomlynSyntax', 'OfficeIMOToml')) {
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        tomldeserialize = [pscustomobject]@{
            Area = 'toml'
            Title = 'TOML model deserialization'
            Libraries = @('OfficeIMO', 'Tomlyn')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Toml.Benchmarks\OfficeIMO.Toml.Benchmarks.csproj'
            Filter = '*TomlDeserializeBenchmarks*'
            ComparisonId = 'toml-deserialize'
            Suite = 'OfficeIMO.Toml.Deserialize'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'CargoWorkspace', 'LockFile')) {
                    foreach ($engine in @('Tomlyn', 'TomlynGenerated', 'OfficeIMOToml', 'OfficeIMOTomlGenerated')) {
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        tomlserialize = [pscustomobject]@{
            Area = 'toml'
            Title = 'TOML model serialization'
            Libraries = @('OfficeIMO', 'Tomlyn')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Toml.Benchmarks\OfficeIMO.Toml.Benchmarks.csproj'
            Filter = '*TomlSerializeBenchmarks*'
            ComparisonId = 'toml-serialize'
            Suite = 'OfficeIMO.Toml.Serialize'
            ExpectedCases = @(
                foreach ($workload in @('SmallConfig', 'CargoWorkspace', 'LockFile')) {
                    foreach ($engine in @('Tomlyn', 'TomlynGenerated', 'OfficeIMOToml', 'OfficeIMOTomlGenerated')) {
                        "$engine|Workload=$workload"
                    }
                }
            )
        }
        imagepngdecode = [pscustomobject]@{
            Area = 'image'
            Title = 'PNG decoding'
            Libraries = @('OfficeIMO', 'SkiaSharp', 'Magick.NET', 'StbImageSharp'; if ($one) { 'OneImage' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImagePngDecodeBenchmarks*'
            ComparisonId = 'image-png-decode'
            Suite = 'OfficeIMO.Drawing.PngDecode'
            ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET', 'StbImageSharp'; if ($one) { 'OneImage' })
        }
        imagejpegdecode = [pscustomobject]@{
            Area = 'image'
            Title = 'JPEG decoding'
            Libraries = @('OfficeIMO', 'SkiaSharp', 'Magick.NET', 'StbImageSharp'; if ($one) { 'OneImage' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImageJpegDecodeBenchmarks*'
            ComparisonId = 'image-jpeg-decode'
            Suite = 'OfficeIMO.Drawing.JpegDecode'
            ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET', 'StbImageSharp'; if ($one) { 'OneImage' })
        }
        imagepngencode = [pscustomobject]@{
            Area = 'image'
            Title = 'PNG encoding'
            Libraries = @('OfficeIMO', 'SkiaSharp', 'Magick.NET'; if ($one) { 'OneImage' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImagePngEncodeBenchmarks*'
            ComparisonId = 'image-png-encode'
            Suite = 'OfficeIMO.Drawing.PngEncode'
            ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET'; if ($one) { 'OneImage' })
        }
        imagetifflzw = [pscustomobject]@{
            Area = 'image'
            Title = 'TIFF LZW round trip'
            Libraries = @('OfficeIMO', 'Magick.NET'; if ($one) { 'OneImage' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImageTiffLzw*Benchmarks*'
            ComparisonId = 'image-tiff-lzw'
            Suite = 'OfficeIMO.Drawing.TiffLzw'
            ExpectedCases = @('OfficeIMOEncode', 'MagickNETEncode', 'OfficeIMODecodeExternal', 'MagickNETDecodeExternal'; if ($one) { 'OneImageEncode'; 'OneImageDecodeExternal' })
        }
        imagewebplossless = [pscustomobject]@{
            Area = 'image'
            Title = 'Lossless WebP encoding and decoding'
            Libraries = @('OfficeIMO', 'Magick.NET'; if ($one) { 'OneImage' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImageWebpLossless*Benchmarks*'
            ComparisonId = 'image-webp-lossless'
            Suite = 'OfficeIMO.Drawing.WebpLossless'
            ExpectedCases = @('OfficeIMOEncode', 'MagickNETEncode', 'OfficeIMODecodeExternal', 'MagickNETDecodeExternal'; if ($one) { 'OneImageEncode'; 'OneImageDecodeExternal' })
        }
        imageresize = [pscustomobject]@{
            Area = 'image'
            Title = 'Image resizing'
            Libraries = @('OfficeIMO', 'SkiaSharp', 'Magick.NET'; if ($one) { 'OneImage' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImageResizeBenchmarks*'
            ComparisonId = 'image-resize'
            Suite = 'OfficeIMO.Drawing.Resize'
            ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET'; if ($one) { 'OneImage' })
        }
        pdfgenerate = [pscustomobject]@{
            Area = 'pdf'
            Title = 'Structured PDF generation'
            Libraries = @('OfficeIMO', 'QuestPDF', 'MigraDoc', 'iText'; if ($one) { 'OnePDF' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfGenerationBenchmarks*'
            ComparisonId = 'pdf-structured-generation'
            Suite = 'OfficeIMO.Pdf.StructuredGeneration'
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High')) {
                    foreach ($engine in @('OfficeIMO', 'QuestPDF', 'MigraDoc', 'IText'; if ($one) { 'OnePdfLayout' })) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        pdfinvoice = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF invoice generation'
            Libraries = @('OfficeIMO', 'QuestPDF', 'iText'; if ($one) { 'OnePDF' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'Existing invoice comparison is not qualified for the public catalog; retain the reviewed exclusion.'
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfInvoiceGenerationBenchmarks*'
            ComparisonId = 'pdf-real-world-invoice'
            Suite = 'OfficeIMO.Pdf.RealWorldInvoice'
            CatalogEligible = $false
            ExpectedCases = @('OfficeIMO', 'QuestPDF', 'IText'; if ($one) { 'OnePdfLayout' })
        }
        pdfinvoiceworkflow = [pscustomobject]@{
            Area = 'pdf'
            Title = 'Typed invoice PDF workflow'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO-only workflow; excluded from public library comparisons.'
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfTypedInvoiceWorkflowBenchmarks*'
            ComparisonId = 'pdf-typed-invoice-workflow'
            Suite = 'OfficeIMO.Pdf.TypedInvoiceWorkflow'
            CatalogEligible = $false
            ExpectedCases = @('TypedElectronicInvoicePdf')
        }
        pdfhtml = [pscustomobject]@{
            Area = 'pdf'
            Title = 'HTML to PDF'
            Libraries = @('OfficeIMO', 'PeachPDF', 'PdfPecker', 'iText pdfHTML', 'Chromium')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfHtmlBenchmarks*'
            ComparisonId = 'pdf-html-generation'
            Suite = 'OfficeIMO.Pdf.HtmlGeneration'
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High')) {
                    foreach ($engine in @('OfficeIMO', 'PeachPDF', 'PdfPecker', 'ITextPdfHtml', 'Chromium')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        pdfhtmlpayload = [pscustomobject]@{
            Area = 'pdf'
            Title = 'HTML 21 KiB payload to PDF'
            Libraries = @('OfficeIMO', 'PeachPDF', 'PdfPecker'; if ($one) { 'OneRender' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfHtmlPayloadBenchmarks*'
            ComparisonId = 'pdf-html-21k-payload'
            Suite = 'OfficeIMO.Pdf.Html21KiBPayload'
            ExpectedCases = @(
                foreach ($payload in @('PlainText', 'Table', 'Multilingual')) {
                    foreach ($engine in @('OfficeIMO', 'PeachPDF', 'PdfPecker'; if ($one) { 'OneRender' })) {
                        "$engine|Payload=$payload"
                    }
                }
            )
        }
        pdfformats = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF format route health'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO-only format route health; excluded from public library comparisons.'
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*Pdf*FormatConversionBenchmarks*'
            ComparisonId = 'officeimo-pdf-format-route-health'
            Suite = 'OfficeIMO.Pdf.FormatConversion'
            CatalogEligible = $false
            ExpectedCases = @(
                foreach ($format in @(
                    'Docx',
                    'Xlsx',
                    'Pptx',
                    'Html',
                    'Markdown',
                    'Rtf',
                    'AsciiDoc',
                    'Latex',
                    'Mhtml',
                    'OneNote',
                    'Odt',
                    'Ods',
                    'Odp',
                    'Visio')) {
                    "ConvertToPdfBytes|Format=$format"
                }
            )
        }
        pdfread = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF text extraction'
            Libraries = @('OfficeIMO', 'PdfPig', 'iText'; if ($one) { 'OnePDF' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = 'OfficeIMO.Pdf.Benchmarks.Comparisons.PdfReadBenchmarks.*'
            ComparisonId = 'pdf-text-extraction'
            Suite = 'OfficeIMO.Pdf.TextExtraction'
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High')) {
                    foreach ($producer in @('OfficeIMO', 'QuestPDF', 'PeachPDF', 'MigraDoc', 'IText')) {
                        foreach ($engine in @('OfficeIMO', 'PdfPig', 'IText'; if ($one) { 'OnePdf'; 'OnePdfParallel' })) {
                            "$engine|Producer=$producer&Scale=$scale"
                        }
                    }
                }
            )
        }
        pdfstructuredread = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF read profile route health'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO-only read profile route health; excluded from public library comparisons.'
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfStructuredRead*Benchmarks*'
            ComparisonId = 'officeimo-pdf-read-profile-route-health'
            Suite = 'OfficeIMO.Pdf.ReadProfileRouteHealth'
            CatalogEligible = $false
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High')) {
                    "FastLoadAndRead|Scale=$scale"
                    "StructuredLoadAndRead|Scale=$scale"
                }
            )
        }
        pdfreverse = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF reverse conversion'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfReverseConversionBenchmarks*'
            ComparisonId = 'pdf-reverse-conversion'
            Suite = 'OfficeIMO.Pdf.ReverseConversion'
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High')) {
                    foreach ($producer in @('OfficeIMO', 'QuestPDF', 'PeachPDF', 'MigraDoc', 'IText')) {
                        foreach ($route in @(
                            'PdfToDocx',
                            'PdfToHtml',
                            'PdfToXlsx',
                            'PdfToPptx',
                            'PdfToOdt',
                            'PdfToOds',
                            'PdfToOdp',
                            'PdfToPng')) {
                            "$route|Producer=$producer&Scale=$scale"
                        }
                    }
                }
            )
        }
        pdfcorpusread = [pscustomobject]@{
            Area = 'pdf'
            Title = 'Independent PDF corpus text extraction'
            Libraries = @('OfficeIMO', 'PdfPig', 'iText'; if ($one) { 'OnePDF' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfCorpusReadBenchmarks*'
            ComparisonId = 'pdf-corpus-text-extraction'
            Suite = 'OfficeIMO.Pdf.CorpusTextExtraction'
            ExpectedCases = @(
                foreach ($document in @('OfficeIMO-500p', 'NIST-492p', 'Type3-85p', 'PDFA-258p-12MB')) {
                    foreach ($engine in @('OfficeIMO', 'PdfPig', 'IText'; if ($one) { 'OnePdf'; 'OnePdfParallel' })) {
                        "$engine|Document=$document"
                    }
                }
            )
        }
        pdfsplit = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF splitting'
            Libraries = @('OfficeIMO', 'iText', 'PDFsharp'; if ($one) { 'OnePDF' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = 'OfficeIMO.Pdf.Benchmarks.Comparisons.PdfSplitBenchmarks.*'
            ComparisonId = 'pdf-split'
            Suite = 'OfficeIMO.Pdf.Split'
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High', 'VeryHigh')) {
                    foreach ($producer in @('OfficeIMO', 'IText')) {
                        foreach ($density in @('Standard', 'Dense')) {
                            foreach ($workflow in @('EveryPage', 'Bundles')) {
                                foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp'; if ($one) { 'OnePdf' })) {
                                    "$engine|Density=$density&Producer=$producer&Scale=$scale&Workflow=$workflow"
                                }
                            }
                        }
                    }
                }
            )
        }
        pdfmerge = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF merging'
            Libraries = @('OfficeIMO', 'iText', 'PDFsharp'; if ($one) { 'OnePDF' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = 'OfficeIMO.Pdf.Benchmarks.Comparisons.PdfMergeBenchmarks.*'
            ComparisonId = 'pdf-merge'
            Suite = 'OfficeIMO.Pdf.Merge'
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High', 'VeryHigh')) {
                    foreach ($producer in @('OfficeIMO', 'IText')) {
                        foreach ($density in @('Standard', 'Dense')) {
                            foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp'; if ($one) { 'OnePdf' })) {
                                "$engine|Density=$density&Producer=$producer&Scale=$scale"
                            }
                        }
                    }
                }
            )
        }
        pdfselect = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF page selection'
            Libraries = @('OfficeIMO', 'iText', 'PDFsharp'; if ($one) { 'OnePDF' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = 'OfficeIMO.Pdf.Benchmarks.Comparisons.PdfPageSelectionBenchmarks.*'
            ComparisonId = 'pdf-page-selection'
            Suite = 'OfficeIMO.Pdf.PageSelection'
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High', 'VeryHigh')) {
                    foreach ($producer in @('OfficeIMO', 'IText')) {
                        foreach ($density in @('Standard', 'Dense')) {
                            foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp'; if ($one) { 'OnePdf' })) {
                                "$engine|Density=$density&Producer=$producer&Scale=$scale"
                            }
                        }
                    }
                }
            )
        }
        wordcreate = [pscustomobject]@{
            Area = 'word'
            Title = 'DOCX paragraph creation'
            Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK'; if ($one) { 'OneDocument.Documents' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
            Filter = '*WordCreateParagraphComparisonBenchmarks*'
            ComparisonId = 'word-docx-create-paragraphs'
            Suite = 'OfficeIMO.Word.CreateParagraphs'
            ExpectedCases = @(
                foreach ($itemCount in @(100, 1000, 10000)) {
                    foreach ($engine in @('OfficeIMO', 'DocX', 'NPOI', 'OpenXmlSdkMinimal', 'OpenXmlSdkEquivalent'; if ($one) { 'OneDocument' })) {
                        "$engine|ItemCount=$itemCount"
                    }
                }
            )
        }
        wordreport = [pscustomobject]@{
            Area = 'word'
            Title = 'DOCX report creation'
            Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK'; if ($one) { 'OneDocument.Documents' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
            Filter = '*WordCreateReportComparisonBenchmarks*'
            ComparisonId = 'word-docx-create-report'
            Suite = 'OfficeIMO.Word.CreateReport'
            ExpectedCases = @(
                'OfficeIMO|RowCount=100'
                'OfficeIMO|RowCount=1000'
                'DocX|RowCount=100'
                'DocX|RowCount=1000'
                'NPOI|RowCount=100'
                'NPOI|RowCount=1000'
                'OpenXmlSdkMinimal|RowCount=100'
                'OpenXmlSdkMinimal|RowCount=1000'
                'OpenXmlSdkEquivalent|RowCount=100'
                'OpenXmlSdkEquivalent|RowCount=1000'
                if ($one) { 'OneDocument|RowCount=100'; 'OneDocument|RowCount=1000' }
            )
        }
        wordread = [pscustomobject]@{
            Area = 'word'
            Title = 'DOCX paragraph reading'
            Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK'; if ($one) { 'OneDocument.Documents' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
            Filter = '*WordReadComparisonBenchmarks*'
            ComparisonId = 'word-docx-read-paragraphs'
            Suite = 'OfficeIMO.Word.ReadParagraphs'
            ExpectedCases = @(
                foreach ($itemCount in @(100, 1000, 10000)) {
                    # OfficeIMOReader* are OfficeIMO.Reader's text projections of the same paragraphs.
                    foreach ($engine in @('OfficeIMO', 'OfficeIMOReaderText', 'OfficeIMOReaderStreamedText', 'DocX', 'NPOI', 'OpenXmlSdk'; if ($one) { 'OneDocument'; 'OneDocumentExtractText' })) {
                        "$engine|ItemCount=$itemCount"
                    }
                }
            )
        }
        wordreplace = [pscustomobject]@{
            Area = 'word'
            Title = 'DOCX text replacement and save'
            Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK'; if ($one) { 'OneDocument.Documents' })
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
            Filter = '*WordReplaceComparisonBenchmarks*'
            ComparisonId = 'word-docx-replace-and-save'
            Suite = 'OfficeIMO.Word.ReplaceAndSave'
            ExpectedCases = @(
                'OfficeIMO|ItemCount=100'
                'OfficeIMO|ItemCount=1000'
                'DocX|ItemCount=100'
                'DocX|ItemCount=1000'
                'NPOI|ItemCount=100'
                'NPOI|ItemCount=1000'
                'OpenXmlSdk|ItemCount=100'
                'OpenXmlSdk|ItemCount=1000'
                if ($one) { 'OneDocument|ItemCount=100'; 'OneDocument|ItemCount=1000' }
            )
        }
        pptxcreate = [pscustomobject]@{
            Area = 'powerpoint'
            Title = 'PPTX create and save'
            Libraries = @('OfficeIMO', 'ShapeCrawler', 'Open XML SDK', 'OneDocument.Presentations')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler/OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler.csproj'
            RunnerArguments = @('bench')
            Frameworks = @('net10.0')
            Filter = '*PowerPointCreateComparisonBenchmarks*'
            ComparisonId = 'pptx-create'
            Suite = 'OfficeIMO.PowerPoint.Create'
            ExpectedCases = @(
                foreach ($slideCount in @(20, 200)) {
                    foreach ($engine in @('OfficeIMO', 'ShapeCrawler', 'OpenXmlSdk', 'OneDocument')) {
                        "$engine|SlideCount=$slideCount"
                    }
                }
            )
        }
        pptxedit = [pscustomobject]@{
            Area = 'powerpoint'
            Title = 'PPTX open, edit, and save'
            Libraries = @('OfficeIMO', 'ShapeCrawler', 'Open XML SDK', 'OneDocument.Presentations')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler/OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler.csproj'
            RunnerArguments = @('bench')
            Frameworks = @('net10.0')
            Filter = '*PowerPointEditComparisonBenchmarks*'
            ComparisonId = 'pptx-edit'
            Suite = 'OfficeIMO.PowerPoint.Edit'
            ExpectedCases = @(
                foreach ($slideCount in @(20, 200)) {
                    foreach ($engine in @('OfficeIMO', 'ShapeCrawler', 'OpenXmlSdk', 'OneDocument')) {
                        "$engine|SlideCount=$slideCount"
                    }
                }
            )
        }

    }

    $areaRoot = Join-Path $PSScriptRoot 'LibraryComparisonDefinitions'
    foreach ($file in @(if (Test-Path -LiteralPath $areaRoot) { Get-ChildItem -LiteralPath $areaRoot -Filter '*.ps1' -File | Sort-Object Name })) {
        $areaDefinitions = & $file.FullName -Framework $Framework
        foreach ($entry in $areaDefinitions.GetEnumerator()) {
            if ($definitions.Contains($entry.Key)) {
                throw "Workload '$($entry.Key)' in $($file.Name) is already declared."
            }
            $definitions[$entry.Key] = $entry.Value
        }
    }

    return $definitions
}
