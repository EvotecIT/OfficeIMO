# Shared workload declarations for collection and site generation. Loading this file performs no work.
function Get-LibraryComparisonDefinitions {
    <#
    .SYNOPSIS
    Returns every library-comparison declaration without executing a runner.
    .PARAMETER Framework
    Target framework included in each comparison identity.
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
            ComparisonId = "csv-million-first-row-$Framework"
            Suite = 'OfficeIMO.CSV.LargeFirstRow'
            CatalogEligible = $false
            FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
            IdentityVariables = @('shape')
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
            ComparisonId = "csv-async-read-$Framework"
            Suite = 'OfficeIMO.CSV.AsyncRead'
            CatalogEligible = $false
            FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
            IdentityVariables = @('operation', 'rowcount', 'shape')
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
            ComparisonId = "csv-wide-async-read-$Framework"
            Suite = 'OfficeIMO.CSV.WideAsyncRead'
            CatalogEligible = $false
            FullRunArguments = @('--warmupCount', '6', '--iterationCount', '12', '--outliers', 'DontRemove')
            IdentityVariables = @('operation', 'rowcount')
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
            ComparisonId = "pdf-native-100-pages-$Framework"
            Suite = 'OfficeIMO.Pdf.NativeOperations'
            IdentityVariables = @('producer', 'operation')
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
            Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sep', 'Sylvan', 'CsvHelper', 'Dataplat.Dbatools.Csv', 'LumenWorks')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
            Filter = '*MarkPflug65KCsvBenchmarks*'
            ComparisonId = "markpflug-65k-csv-decoded-$Framework"
            Suite = 'OfficeIMO.CSV.MarkPflug65K'
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMO', 'ExcelReaderNet', 'Sep', 'Sylvan', 'CsvHelper', 'DataplatDbatools', 'LumenWorks')
        }
        csvwrite = [pscustomobject]@{
            Area = 'csv'
            Title = 'CSV data-reader writing'
            Libraries = @('OfficeIMO', 'Sylvan')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
            Filter = '*CsvDataReaderWriteBenchmarks*(RowCount: 25000,*'
            ComparisonId = "csv-25k-datareader-write-$Framework"
            Suite = 'OfficeIMO.CSV.DataReaderWrite25K'
            IdentityVariables = @('rowcount', 'shape')
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
            Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sylvan', 'ExcelDataReader')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
            Filter = '*MarkPflug65KXlsBenchmarks*'
            ComparisonId = "markpflug-65k-xls-typed-$Framework"
            Suite = 'OfficeIMO.Excel.Xls.MarkPflug65K'
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMO', 'ExcelReaderNet', 'Sylvan', 'ExcelDataReader')
        }
        xlsx = [pscustomobject]@{
            Area = 'excel'
            Title = 'XLSX 65K typed reading'
            Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sylvan', 'ExcelDataReader', 'ClosedXML', 'EPPlus', 'MiniExcel', 'Excel-PRIME', 'HypeLab.IO.Excel', 'XLibur')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
            Filter = '*MarkPflug65KXlsxBenchmarks*'
            ComparisonId = "markpflug-65k-xlsx-typed-$Framework"
            Suite = 'OfficeIMO.Excel.Xlsx.MarkPflug65K'
            IdentityVariables = @()
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
                # These factories/options are compiled only for the modern ExcelReader.NET API.
                if ($Framework -eq 'net10.0') {
                    'ExcelReaderNet_Memory'
                    'ExcelReaderNet_InternStrings'
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
            Libraries = @('OfficeIMO', 'SpreadCheetah', 'Sylvan', 'LargeXlsx')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
            Filter = '*ExcelDataReaderWriteBenchmarks*'
            ComparisonId = "xlsx-25k-datareader-write-$Framework"
            Suite = 'OfficeIMO.Excel.DataReaderWrite25K'
            IdentityVariables = @('rowcount')
            ExpectedCases = @(
                'OfficeIMO|RowCount=25000'
                'SpreadCheetah|RowCount=25000'
                'Sylvan|RowCount=25000'
                'LargeXlsx|RowCount=25000'
            )
        }
        xlsb = [pscustomobject]@{
            Area = 'excel'
            Title = 'XLSB 65K typed reading'
            Libraries = @('OfficeIMO', 'ExcelReader.NET', 'Sylvan', 'ExcelDataReader', 'Excel-PRIME')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
            Filter = '*MarkPflug65KXlsbBenchmarks*'
            ComparisonId = "markpflug-65k-xlsb-typed-$Framework"
            Suite = 'OfficeIMO.Excel.Xlsb.MarkPflug65K'
            IdentityVariables = @()
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
            ComparisonId = "markdown-commonmark-parse-$Framework"
            Suite = 'OfficeIMO.Markdown.CommonMarkParse'
            IdentityVariables = @('corpusname')
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
            ComparisonId = "markdown-commonmark-html-$Framework"
            Suite = 'OfficeIMO.Markdown.CommonMarkHtml'
            IdentityVariables = @('corpusname')
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
            ComparisonId = "html-to-markdown-$Framework"
            Suite = 'OfficeIMO.Markdown.HtmlToMarkdown'
            IdentityVariables = @('corpusname')
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
            ComparisonId = "rtf-to-html-$Framework"
            Suite = 'OfficeIMO.Rtf.HtmlComparison'
            ValidatedSizeEvidence = $true
            IdentityVariables = @('scale')
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
            ComparisonId = "opendocument-ods-create-$Framework"
            Suite = 'OfficeIMO.OpenDocument.OdsCreate'
            ValidatedSizeEvidence = $true
            IdentityVariables = @('scale')
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
            ComparisonId = "opendocument-ods-read-$Framework"
            Suite = 'OfficeIMO.OpenDocument.OdsRead'
            IdentityVariables = @('scale')
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
            Libraries = @('OfficeIMO', 'MimeKit')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Email.Benchmarks.Comparisons\OfficeIMO.Email.Benchmarks.Comparisons.csproj'
            Filter = '*EmailMimeReadBenchmarks*'
            ComparisonId = "email-mime-read-$Framework"
            Suite = 'OfficeIMO.Email.MimeRead'
            ValidatedSizeEvidence = $true
            IdentityVariables = @('scale')
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal')) {
                    foreach ($engine in @('OfficeIMO', 'MimeKit')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        emailmimewrite = [pscustomobject]@{
            Area = 'email'
            Title = 'MIME email writing'
            Libraries = @('OfficeIMO', 'MimeKit')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Email.Benchmarks.Comparisons\OfficeIMO.Email.Benchmarks.Comparisons.csproj'
            Filter = '*EmailMimeWriteBenchmarks*'
            ComparisonId = "email-mime-write-$Framework"
            Suite = 'OfficeIMO.Email.MimeWrite'
            ValidatedSizeEvidence = $true
            IdentityVariables = @('scale')
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal')) {
                    foreach ($engine in @('OfficeIMO', 'MimeKit')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        epubread = [pscustomobject]@{
            Area = 'epub'
            Title = 'EPUB opening and reading'
            Libraries = @('OfficeIMO', 'VersOne.Epub', 'HtmlAgilityPack')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Epub.Benchmarks.Comparisons\OfficeIMO.Epub.Benchmarks.Comparisons.csproj'
            Filter = '*EpubReadComparisonBenchmarks*'
            ComparisonId = "epub-open-read-$Framework"
            Suite = 'OfficeIMO.Epub.OpenRead'
            ValidatedSizeEvidence = $true
            IdentityVariables = @('scale')
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
            Libraries = @('OfficeIMO', 'System.IO.Compression')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO enforces path and expansion safety; the platform lane projects raw metadata without equivalent safety policy. Not a public like-for-like safety comparison.'
            Project = 'OfficeIMO.Archives.Benchmarks\OfficeIMO.Archives.Benchmarks.csproj'
            Filter = '*ZipTraversalComparisonBenchmarks*'
            ComparisonId = "zip-safe-traversal-$Framework"
            Suite = 'OfficeIMO.Archives.SafeZipTraversal'
            CatalogEligible = $false
            CollectNonPublishedEvidence = $true
            IdentityVariables = @('scale')
            ExpectedCases = @(
                foreach ($scale in @('Small', 'Normal', 'Large')) {
                    foreach ($engine in @('OfficeIMO', 'Platform')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        imagepngdecode = [pscustomobject]@{
            Area = 'image'
            Title = 'PNG decoding'
            Libraries = @('OfficeIMO', 'SkiaSharp', 'Magick.NET', 'StbImageSharp')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImagePngDecodeBenchmarks*'
            ComparisonId = "image-png-decode-$Framework"
            Suite = 'OfficeIMO.Drawing.PngDecode'
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET', 'StbImageSharp')
        }
        imagejpegdecode = [pscustomobject]@{
            Area = 'image'
            Title = 'JPEG decoding'
            Libraries = @('OfficeIMO', 'SkiaSharp', 'Magick.NET', 'StbImageSharp')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImageJpegDecodeBenchmarks*'
            ComparisonId = "image-jpeg-decode-$Framework"
            Suite = 'OfficeIMO.Drawing.JpegDecode'
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET', 'StbImageSharp')
        }
        imagepngencode = [pscustomobject]@{
            Area = 'image'
            Title = 'PNG encoding'
            Libraries = @('OfficeIMO', 'SkiaSharp', 'Magick.NET')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImagePngEncodeBenchmarks*'
            ComparisonId = "image-png-encode-$Framework"
            Suite = 'OfficeIMO.Drawing.PngEncode'
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET')
        }
        imagetifflzw = [pscustomobject]@{
            Area = 'image'
            Title = 'TIFF LZW round trip'
            Libraries = @('OfficeIMO', 'Magick.NET')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImageTiffLzw*Benchmarks*'
            ComparisonId = "image-tiff-lzw-$Framework"
            Suite = 'OfficeIMO.Drawing.TiffLzw'
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMOEncode', 'MagickNETEncode', 'OfficeIMODecodeExternal', 'MagickNETDecodeExternal')
        }
        imagewebplossless = [pscustomobject]@{
            Area = 'image'
            Title = 'Lossless WebP encoding and decoding'
            Libraries = @('OfficeIMO', 'Magick.NET')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImageWebpLossless*Benchmarks*'
            ComparisonId = "image-webp-lossless-$Framework"
            Suite = 'OfficeIMO.Drawing.WebpLossless'
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMOEncode', 'MagickNETEncode', 'OfficeIMODecodeExternal', 'MagickNETDecodeExternal')
        }
        imageresize = [pscustomobject]@{
            Area = 'image'
            Title = 'Image resizing'
            Libraries = @('OfficeIMO', 'SkiaSharp', 'Magick.NET')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
            Filter = '*ImageResizeBenchmarks*'
            ComparisonId = "image-resize-$Framework"
            Suite = 'OfficeIMO.Drawing.Resize'
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET')
        }
        pdfgenerate = [pscustomobject]@{
            Area = 'pdf'
            Title = 'Structured PDF generation'
            Libraries = @('OfficeIMO', 'QuestPDF', 'MigraDoc', 'iText')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'QuestPDF is pinned to MIT-compatible 2026.5.0; alternate versions require confirmed authorization and internal-only output.'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfGenerationBenchmarks*'
            ComparisonId = "pdf-structured-generation-$Framework"
            Suite = 'OfficeIMO.Pdf.StructuredGeneration'
            IdentityVariables = @('scale')
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High')) {
                    foreach ($engine in @('OfficeIMO', 'QuestPDF', 'MigraDoc', 'IText')) {
                        "$engine|Scale=$scale"
                    }
                }
            )
        }
        pdfinvoice = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF invoice generation'
            Libraries = @('OfficeIMO', 'QuestPDF', 'iText')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'Existing invoice comparison is not qualified for the public catalog; retain the reviewed exclusion.'
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfInvoiceGenerationBenchmarks*'
            ComparisonId = "pdf-real-world-invoice-$Framework"
            Suite = 'OfficeIMO.Pdf.RealWorldInvoice'
            CatalogEligible = $false
            IdentityVariables = @()
            ExpectedCases = @('OfficeIMO', 'QuestPDF', 'IText')
        }
        pdfinvoiceworkflow = [pscustomobject]@{
            Area = 'pdf'
            Title = 'Typed invoice PDF workflow'
            Libraries = @('OfficeIMO')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'OfficeIMO-only workflow; excluded from public library comparisons.'
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfTypedInvoiceWorkflowBenchmarks*'
            ComparisonId = "pdf-typed-invoice-workflow-$Framework"
            Suite = 'OfficeIMO.Pdf.TypedInvoiceWorkflow'
            CatalogEligible = $false
            IdentityVariables = @()
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
            ComparisonId = "pdf-html-generation-$Framework"
            Suite = 'OfficeIMO.Pdf.HtmlGeneration'
            IdentityVariables = @('scale')
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
            Libraries = @('OfficeIMO', 'PeachPDF', 'PdfPecker')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfHtmlPayloadBenchmarks*'
            ComparisonId = "pdf-html-21k-payload-$Framework"
            Suite = 'OfficeIMO.Pdf.Html21KiBPayload'
            IdentityVariables = @('payload')
            ExpectedCases = @(
                foreach ($payload in @('PlainText', 'Table', 'Multilingual')) {
                    foreach ($engine in @('OfficeIMO', 'PeachPDF', 'PdfPecker')) {
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
            ComparisonId = "officeimo-pdf-format-route-health-$Framework"
            Suite = 'OfficeIMO.Pdf.FormatConversion'
            CatalogEligible = $false
            IdentityVariables = @('format')
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
            Libraries = @('OfficeIMO', 'PdfPig', 'iText')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'QuestPDF is pinned to MIT-compatible 2026.5.0; alternate versions require confirmed authorization and internal-only output.'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfReadBenchmarks*'
            ComparisonId = "pdf-text-extraction-$Framework"
            Suite = 'OfficeIMO.Pdf.TextExtraction'
            IdentityVariables = @('producer', 'scale')
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High')) {
                    foreach ($producer in @('OfficeIMO', 'QuestPDF', 'PeachPDF', 'MigraDoc', 'IText')) {
                        foreach ($engine in @('OfficeIMO', 'PdfPig', 'IText')) {
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
            ComparisonId = "officeimo-pdf-read-profile-route-health-$Framework"
            Suite = 'OfficeIMO.Pdf.ReadProfileRouteHealth'
            CatalogEligible = $false
            IdentityVariables = @('scale')
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
            PublishRestriction = 'QuestPDF is pinned to MIT-compatible 2026.5.0; alternate versions require confirmed authorization and internal-only output.'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfReverseConversionBenchmarks*'
            ComparisonId = "pdf-reverse-conversion-$Framework"
            Suite = 'OfficeIMO.Pdf.ReverseConversion'
            IdentityVariables = @('producer', 'scale')
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
            Libraries = @('OfficeIMO', 'PdfPig', 'iText')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfCorpusReadBenchmarks*'
            ComparisonId = "pdf-corpus-text-extraction-$Framework"
            Suite = 'OfficeIMO.Pdf.CorpusTextExtraction'
            IdentityVariables = @('document')
            ExpectedCases = @(
                foreach ($document in @('OfficeIMO-500p', 'NIST-492p', 'Type3-85p', 'PDFA-258p-12MB')) {
                    foreach ($engine in @('OfficeIMO', 'PdfPig', 'IText')) {
                        "$engine|Document=$document"
                    }
                }
            )
        }
        pdfsplit = [pscustomobject]@{
            Area = 'pdf'
            Title = 'PDF splitting'
            Libraries = @('OfficeIMO', 'iText', 'PDFsharp')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfSplitBenchmarks*'
            ComparisonId = "pdf-split-$Framework"
            Suite = 'OfficeIMO.Pdf.Split'
            IdentityVariables = @('density', 'producer', 'scale', 'workflow')
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High', 'VeryHigh')) {
                    foreach ($producer in @('OfficeIMO', 'IText')) {
                        foreach ($density in @('Standard', 'Dense')) {
                            foreach ($workflow in @('EveryPage', 'Bundles')) {
                                foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp')) {
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
            Libraries = @('OfficeIMO', 'iText', 'PDFsharp')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfMergeBenchmarks*'
            ComparisonId = "pdf-merge-$Framework"
            Suite = 'OfficeIMO.Pdf.Merge'
            IdentityVariables = @('density', 'producer', 'scale')
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High', 'VeryHigh')) {
                    foreach ($producer in @('OfficeIMO', 'IText')) {
                        foreach ($density in @('Standard', 'Dense')) {
                            foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp')) {
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
            Libraries = @('OfficeIMO', 'iText', 'PDFsharp')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
            Filter = '*PdfPageSelectionBenchmarks*'
            ComparisonId = "pdf-page-selection-$Framework"
            Suite = 'OfficeIMO.Pdf.PageSelection'
            IdentityVariables = @('density', 'producer', 'scale')
            ExpectedCases = @(
                foreach ($scale in @('Easy', 'Medium', 'High', 'VeryHigh')) {
                    foreach ($producer in @('OfficeIMO', 'IText')) {
                        foreach ($density in @('Standard', 'Dense')) {
                            foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp')) {
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
            Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'DocX Xceed Community License prohibits benchmark publication without advance permission; NPOI 2.8.1 requires explicit OSMF EULA acknowledgement for execution.'
            CatalogEligible = $false
            Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
            Filter = '*WordCreateParagraphComparisonBenchmarks*'
            ComparisonId = "word-docx-create-paragraphs-$Framework"
            Suite = 'OfficeIMO.Word.CreateParagraphs'
            IdentityVariables = @('itemcount')
            ExpectedCases = @(
                foreach ($itemCount in @(100, 1000, 10000)) {
                    foreach ($engine in @('OfficeIMO', 'DocX', 'NPOI', 'OpenXmlSdk')) {
                        "$engine|ItemCount=$itemCount"
                    }
                }
            )
        }
        wordreport = [pscustomobject]@{
            Area = 'word'
            Title = 'DOCX report creation'
            Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'DocX Xceed Community License prohibits benchmark publication without advance permission; NPOI 2.8.1 requires explicit OSMF EULA acknowledgement for execution.'
            CatalogEligible = $false
            Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
            Filter = '*WordCreateReportComparisonBenchmarks*'
            ComparisonId = "word-docx-create-report-$Framework"
            Suite = 'OfficeIMO.Word.CreateReport'
            IdentityVariables = @('rowcount')
            ExpectedCases = @(
                'OfficeIMO|RowCount=100'
                'OfficeIMO|RowCount=1000'
                'DocX|RowCount=100'
                'DocX|RowCount=1000'
                'NPOI|RowCount=100'
                'NPOI|RowCount=1000'
                'OpenXmlSdk|RowCount=100'
                'OpenXmlSdk|RowCount=1000'
            )
        }
        wordread = [pscustomobject]@{
            Area = 'word'
            Title = 'DOCX paragraph reading'
            Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'DocX Xceed Community License prohibits benchmark publication without advance permission; NPOI 2.8.1 requires explicit OSMF EULA acknowledgement for execution.'
            CatalogEligible = $false
            Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
            Filter = '*WordReadComparisonBenchmarks*'
            ComparisonId = "word-docx-read-paragraphs-$Framework"
            Suite = 'OfficeIMO.Word.ReadParagraphs'
            IdentityVariables = @('itemcount')
            ExpectedCases = @(
                foreach ($itemCount in @(100, 1000, 10000)) {
                    foreach ($engine in @('OfficeIMO', 'DocX', 'NPOI', 'OpenXmlSdk')) {
                        "$engine|ItemCount=$itemCount"
                    }
                }
            )
        }
        wordreplace = [pscustomobject]@{
            Area = 'word'
            Title = 'DOCX text replacement and save'
            Libraries = @('OfficeIMO', 'DocX', 'NPOI', 'Open XML SDK')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'DocX Xceed Community License prohibits benchmark publication without advance permission; NPOI 2.8.1 requires explicit OSMF EULA acknowledgement for execution.'
            CatalogEligible = $false
            Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
            Filter = '*WordReplaceComparisonBenchmarks*'
            ComparisonId = "word-docx-replace-and-save-$Framework"
            Suite = 'OfficeIMO.Word.ReplaceAndSave'
            IdentityVariables = @('itemcount')
            ExpectedCases = @(
                'OfficeIMO|ItemCount=100'
                'OfficeIMO|ItemCount=1000'
                'DocX|ItemCount=100'
                'DocX|ItemCount=1000'
                'NPOI|ItemCount=100'
                'NPOI|ItemCount=1000'
                'OpenXmlSdk|ItemCount=100'
                'OpenXmlSdk|ItemCount=1000'
            )
        }
        pptxcreate = [pscustomobject]@{
            Area = 'powerpoint'
            Title = 'PPTX create and save'
            Libraries = @('OfficeIMO', 'ShapeCrawler', 'Open XML SDK')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler/OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler.csproj'
            RunnerArguments = @('bench')
            Frameworks = @('net10.0')
            Filter = '*PowerPointCreateComparisonBenchmarks*'
            ComparisonId = "pptx-create-$Framework"
            Suite = 'OfficeIMO.PowerPoint.Create'
            IdentityVariables = @('slidecount')
            ExpectedCases = @(
                foreach ($slideCount in @(20, 200)) {
                    foreach ($engine in @('OfficeIMO', 'ShapeCrawler', 'OpenXmlSdk')) {
                        "$engine|SlideCount=$slideCount"
                    }
                }
            )
        }
        pptxedit = [pscustomobject]@{
            Area = 'powerpoint'
            Title = 'PPTX open, edit, and save'
            Libraries = @('OfficeIMO', 'ShapeCrawler', 'Open XML SDK')
            ExpectedPlatforms = @('windows', 'linux', 'macos')
            PublishRestriction = 'None'
            CatalogEligible = $true
            Project = 'OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler/OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler.csproj'
            RunnerArguments = @('bench')
            Frameworks = @('net10.0')
            Filter = '*PowerPointEditComparisonBenchmarks*'
            ComparisonId = "pptx-edit-$Framework"
            Suite = 'OfficeIMO.PowerPoint.Edit'
            IdentityVariables = @('slidecount')
            ExpectedCases = @(
                foreach ($slideCount in @(20, 200)) {
                    foreach ($engine in @('OfficeIMO', 'ShapeCrawler', 'OpenXmlSdk')) {
                        "$engine|SlideCount=$slideCount"
                    }
                }
            )
        }

    }

    return $definitions
}
