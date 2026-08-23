param(
    [ValidateSet('quick', 'full')]
    [string] $RunMode = 'quick',
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [ValidateSet(
        'all',
        'csv',
        'csvwrite',
        'xls',
        'xlsx',
        'xlsxwrite',
        'xlsb',
        'markdownparse',
        'markdownhtml',
        'htmltomarkdown',
        'rtfhtml',
        'odscreate',
        'odsread',
        'imagepngdecode',
        'imagejpegdecode',
        'imagepngencode',
        'imageresize',
        'word',
        'wordcreate',
        'wordreport',
        'wordread',
        'wordreplace',
        'pdfgenerate',
        'pdfhtml',
        'pdfhtmlpayload',
        'pdfformats',
        'pdfread',
        'pdfreverse',
        'pdfcorpusread',
        'pdfsplit',
        'pdfmerge',
        'pdfselect')]
    [string] $Workload = 'all',
    [string] $PdfCorpusRoot,
    [string] $OutputRoot = (Join-Path ([System.IO.Path]::GetTempPath()) 'OfficeIMO\Benchmarks\Runs'),
    [string] $PowerForgeRoot = $env:POWERFORGE_ROOT,
    [ValidateSet('net8.0', 'net10.0')]
    [string] $PowerForgeFramework = 'net8.0',
    [UInt64] $AffinityMask = 0,
    [switch] $AcceptNPOIOSMFLicense,
    [switch] $Publish,
    [switch] $PlanOnly
)

$ErrorActionPreference = 'Stop'

if ($Publish -and $RunMode -ne 'full') {
    throw 'Only a full BenchmarkDotNet run can be marked publishable.'
}

. (Join-Path $PSScriptRoot 'BenchmarkEvidence.ps1')

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$affinityLabel = if ($AffinityMask -ne 0) { '0x{0:X}' -f $AffinityMask } else { $null }
$OutputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
    $OutputRoot)
$platform = if ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Windows)) {
    'windows'
} elseif ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Linux)) {
    'linux'
} elseif ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::OSX)) {
    'macos'
} else {
    throw 'Unsupported benchmark platform.'
}

$definitions = [ordered]@{
    csv = [pscustomobject]@{
        Project = 'OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj'
        Filter = '*MarkPflug65KCsvBenchmarks*'
        ComparisonId = "markpflug-65k-csv-decoded-$Framework"
        Suite = 'OfficeIMO.CSV.MarkPflug65K'
        IdentityVariables = @()
        ExpectedCases = @('OfficeIMO', 'Sep', 'Sylvan', 'CsvHelper', 'DataplatDbatools', 'LumenWorks')
    }
    csvwrite = [pscustomobject]@{
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
        Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
        Filter = '*MarkPflug65KXlsBenchmarks*'
        ComparisonId = "markpflug-65k-xls-typed-$Framework"
        Suite = 'OfficeIMO.Excel.Xls.MarkPflug65K'
        IdentityVariables = @()
        ExpectedCases = @('OfficeIMO', 'Sylvan', 'ExcelDataReader')
    }
    xlsx = [pscustomobject]@{
        Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
        Filter = '*MarkPflug65KXlsxBenchmarks*'
        ComparisonId = "markpflug-65k-xlsx-typed-$Framework"
        Suite = 'OfficeIMO.Excel.Xlsx.MarkPflug65K'
        IdentityVariables = @()
        ExpectedCases = @('OfficeIMO', 'Sylvan', 'ExcelDataReader', 'ClosedXML', 'EPPlus', 'MiniExcel')
    }
    xlsxwrite = [pscustomobject]@{
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
        Project = 'OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj'
        Filter = '*MarkPflug65KXlsbBenchmarks*'
        ComparisonId = "markpflug-65k-xlsb-typed-$Framework"
        Suite = 'OfficeIMO.Excel.Xlsb.MarkPflug65K'
        IdentityVariables = @()
        ExpectedCases = @('OfficeIMO', 'Sylvan', 'ExcelDataReader')
    }
    markdownparse = [pscustomobject]@{
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
                foreach ($engine in @('OfficeIMO-Semantic', 'Markdig')) {
                    "$engine|CorpusName=$corpus"
                }
            }
        )
    }
    markdownhtml = [pscustomobject]@{
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
                foreach ($engine in @('OfficeIMO', 'Markdig')) {
                    "$engine|CorpusName=$corpus"
                }
            }
        )
    }
    htmltomarkdown = [pscustomobject]@{
        Project = 'OfficeIMO.Markdown.Benchmarks\OfficeIMO.Markdown.Benchmarks.csproj'
        Filter = '*HtmlToMarkdownBenchmarks*'
        ComparisonId = "html-to-markdown-$Framework"
        Suite = 'OfficeIMO.Markdown.HtmlToMarkdown'
        IdentityVariables = @('corpusname')
        ExpectedCases = @(
            foreach ($corpus in @('Article', 'LargeArticle', 'Table')) {
                foreach ($engine in @('OfficeIMO', 'ReverseMarkdown')) {
                    "$engine|CorpusName=$corpus"
                }
            }
        )
    }
    rtfhtml = [pscustomobject]@{
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
    imagepngdecode = [pscustomobject]@{
        Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
        Filter = '*ImagePngDecodeBenchmarks*'
        ComparisonId = "image-png-decode-$Framework"
        Suite = 'OfficeIMO.Drawing.PngDecode'
        IdentityVariables = @()
        ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET', 'StbImageSharp')
    }
    imagejpegdecode = [pscustomobject]@{
        Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
        Filter = '*ImageJpegDecodeBenchmarks*'
        ComparisonId = "image-jpeg-decode-$Framework"
        Suite = 'OfficeIMO.Drawing.JpegDecode'
        IdentityVariables = @()
        ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET', 'StbImageSharp')
    }
    imagepngencode = [pscustomobject]@{
        Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
        Filter = '*ImagePngEncodeBenchmarks*'
        ComparisonId = "image-png-encode-$Framework"
        Suite = 'OfficeIMO.Drawing.PngEncode'
        IdentityVariables = @()
        ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET')
    }
    imageresize = [pscustomobject]@{
        Project = 'OfficeIMO.Drawing.Benchmarks.Comparisons\OfficeIMO.Drawing.Benchmarks.Comparisons.csproj'
        Filter = '*ImageResizeBenchmarks*'
        ComparisonId = "image-resize-$Framework"
        Suite = 'OfficeIMO.Drawing.Resize'
        IdentityVariables = @()
        ExpectedCases = @('OfficeIMO', 'SkiaSharp', 'MagickNET')
    }
    pdfgenerate = [pscustomobject]@{
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
    pdfhtml = [pscustomobject]@{
        Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
        Filter = '*PdfHtmlBenchmarks*'
        ComparisonId = "pdf-html-generation-$Framework"
        Suite = 'OfficeIMO.Pdf.HtmlGeneration'
        IdentityVariables = @('scale')
        ExpectedCases = @(
            foreach ($scale in @('Easy', 'Medium', 'High')) {
                foreach ($engine in @('OfficeIMO', 'PeachPDF')) {
                    "$engine|Scale=$scale"
                }
            }
        )
    }
    pdfhtmlpayload = [pscustomobject]@{
        Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
        Filter = '*PdfHtmlPayloadBenchmarks*'
        ComparisonId = "pdf-html-21k-payload-$Framework"
        Suite = 'OfficeIMO.Pdf.Html21KiBPayload'
        IdentityVariables = @('payload')
        ExpectedCases = @(
            foreach ($payload in @('PlainText', 'Table', 'Multilingual')) {
                foreach ($engine in @('OfficeIMO', 'PeachPDF')) {
                    "$engine|Payload=$payload"
                }
            }
        )
    }
    pdfformats = [pscustomobject]@{
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
                "ConvertToPdf|Format=$format"
            }
        )
    }
    pdfread = [pscustomobject]@{
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
    pdfreverse = [pscustomobject]@{
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
        Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
        Filter = '*PdfCorpusReadBenchmarks*'
        ComparisonId = "pdf-corpus-text-extraction-$Framework"
        Suite = 'OfficeIMO.Pdf.CorpusTextExtraction'
        Default = $false
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
        Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
        Filter = '*PdfSplitBenchmarks*'
        ComparisonId = "pdf-split-$Framework"
        Suite = 'OfficeIMO.Pdf.Split'
        IdentityVariables = @('producer', 'scale', 'workflow')
        ExpectedCases = @(
            foreach ($scale in @('Easy', 'Medium', 'High')) {
                foreach ($producer in @('OfficeIMO', 'IText')) {
                    foreach ($workflow in @('EveryPage', 'Bundles')) {
                        foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp')) {
                            "$engine|Producer=$producer&Scale=$scale&Workflow=$workflow"
                        }
                    }
                }
            }
        )
    }
    pdfmerge = [pscustomobject]@{
        Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
        Filter = '*PdfMergeBenchmarks*'
        ComparisonId = "pdf-merge-$Framework"
        Suite = 'OfficeIMO.Pdf.Merge'
        IdentityVariables = @('producer', 'scale')
        ExpectedCases = @(
            foreach ($scale in @('Easy', 'Medium', 'High')) {
                foreach ($producer in @('OfficeIMO', 'IText')) {
                    foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp')) {
                        "$engine|Producer=$producer&Scale=$scale"
                    }
                }
            }
        )
    }
    pdfselect = [pscustomobject]@{
        Project = 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
        Filter = '*PdfPageSelectionBenchmarks*'
        ComparisonId = "pdf-page-selection-$Framework"
        Suite = 'OfficeIMO.Pdf.PageSelection'
        IdentityVariables = @('producer', 'scale')
        ExpectedCases = @(
            foreach ($scale in @('Easy', 'Medium', 'High')) {
                foreach ($producer in @('OfficeIMO', 'IText')) {
                    foreach ($engine in @('OfficeIMO', 'IText', 'PdfSharp')) {
                        "$engine|Producer=$producer&Scale=$scale"
                    }
                }
            }
        )
    }
    wordcreate = [pscustomobject]@{
        Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
        Filter = '*WordCreateParagraphComparisonBenchmarks*'
        ComparisonId = "word-docx-create-paragraphs-$Framework"
        Suite = 'OfficeIMO.Word.CreateParagraphs'
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
    wordreport = [pscustomobject]@{
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
        Project = 'OfficeIMO.Word.Benchmarks\OfficeIMO.Word.Benchmarks.csproj'
        Filter = '*WordReadComparisonBenchmarks*'
        ComparisonId = "word-docx-read-paragraphs-$Framework"
        Suite = 'OfficeIMO.Word.ReadParagraphs'
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
    wordreplace = [pscustomobject]@{
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
}

$selected = if ($Workload -eq 'all') {
    @(
        $definitions.GetEnumerator() |
            Where-Object { $_.Key -notlike 'word*' -and $_.Value.Default -ne $false } |
            ForEach-Object { $_.Key }
    )
} elseif ($Workload -eq 'word') {
    @('wordcreate', 'wordreport', 'wordread', 'wordreplace')
} else {
    @($Workload)
}

if ($Workload -eq 'pdfcorpusread' -and [string]::IsNullOrWhiteSpace($PdfCorpusRoot)) {
    throw 'PdfCorpusRoot is required for the pdfcorpusread workload. Prepare the corpus first.'
}

$containsWordWorkload = @($selected | Where-Object { $_ -like 'word*' }).Count -gt 0

if ($Publish -and $containsWordWorkload) {
    throw @'
DocX is distributed under the Xceed Community License, which prohibits publishing benchmark or performance comparison results without Xceed's advance permission. Word comparison evidence is local-only. Obtain written permission and update this reviewed publication gate before publishing it.
'@
}

if ($containsWordWorkload -and -not $AcceptNPOIOSMFLicense) {
    throw @'
The Word comparison suite includes NPOI 2.8.0. Review the NPOI binary EULA at https://github.com/nissl-lab/npoi/blob/master/OSMFEULA.txt, then rerun with -AcceptNPOIOSMFLicense to acknowledge it for this opt-in benchmark run.
'@
}

$executionPlan = @(
    foreach ($name in $selected) {
        $definition = $definitions[$name]
        $catalogEligibleByPolicy = $name -notlike 'word*' -and
            ($null -eq $definition.PSObject.Properties['CatalogEligible'] -or [bool] $definition.CatalogEligible)
        $willCatalog = ($RunMode -eq 'quick' -or [bool] $Publish) -and $catalogEligibleByPolicy
        [pscustomobject]@{
            Workload = $name
            ComparisonId = $definition.ComparisonId
            CatalogEligible = $catalogEligibleByPolicy
            WillCatalog = $willCatalog
            Publish = [bool] $Publish -and $willCatalog
        }
    }
)
$executionPlanByWorkload = @{}
foreach ($item in $executionPlan) {
    $executionPlanByWorkload[$item.Workload] = $item
}

if ($PlanOnly) {
    $executionPlan
    return
}

if ([string]::IsNullOrWhiteSpace($PowerForgeRoot)) {
    Import-Module PSPublishModule -MinimumVersion 3.0.84 -Force
} else {
    $powerForgeModule = Join-Path $PowerForgeRoot "PSPublishModule\bin\Release\$PowerForgeFramework\PSPublishModule.dll"
    if (-not (Test-Path -LiteralPath $powerForgeModule -PathType Leaf)) {
        throw "The local PowerForge binary was not found at '$powerForgeModule'. Build PSPublishModule for $PowerForgeFramework in Release configuration first."
    }
    Import-Module $powerForgeModule -Force
}

$stamp = [DateTimeOffset]::UtcNow.ToString('yyyyMMdd-HHmmss')
$staticRoot = Join-Path $repositoryRoot 'Website\static\data\benchmarks\library-comparisons'
$catalogPath = Join-Path $staticRoot 'index.json'
$catalogEligible = @($executionPlan | Where-Object WillCatalog).Count -gt 0
New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null
if ($catalogEligible) {
    New-Item -ItemType Directory -Force -Path $staticRoot | Out-Null
}

$gitSha = (& git -C $repositoryRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($gitSha)) {
    throw 'Unable to resolve the OfficeIMO source commit for benchmark provenance.'
}
$gitDirty = @(& git -C $repositoryRoot status --porcelain --untracked-files=normal).Count -gt 0
if ($catalogEligible -and $gitDirty) {
    throw 'Cataloged benchmark evidence requires a clean Git worktree so the recorded source commit identifies the measured code exactly.'
}

$measurements = [System.Collections.Generic.List[object]]::new()
foreach ($name in $selected) {
    $definition = $definitions[$name]
    $workloadPlan = $executionPlanByWorkload[$name]
    $artifactsPath = Join-Path $OutputRoot "$platform-$name-$RunMode-$stamp"
    New-Item -ItemType Directory -Force -Path $artifactsPath | Out-Null
    $provenanceMetadata = [ordered]@{
        'benchmark.workload.id' = $definition.ComparisonId
        'benchmark.workload.sourceCommit' = $gitSha
        'benchmark.workload.framework' = $Framework
    }
    if ($definition.ValidatedSizeEvidence) {
        $provenanceMetadata['benchmark.workload.sizeEvidencePath'] = 'validated-size-evidence.json'
    }
    if ($null -ne $affinityLabel) {
        $provenanceMetadata['benchmark.workload.affinityMask'] = $affinityLabel
    }
    $provenanceCapture = Start-BenchmarkProvenanceCapture `
        -SourceRoot $repositoryRoot `
        -ArtifactRoot $artifactsPath `
        -Metadata $provenanceMetadata `
        -RunMode $RunMode

    $arguments = @(
        'run',
        '-c', 'Release',
        '-f', $Framework,
        '--project', (Join-Path $repositoryRoot $definition.Project)
    )
    if ($name -like 'word*') {
        $arguments += '-p:AcceptNPOIOSMFLicense=true'
    }
    $arguments += @(
        '--',
        '--filter', $definition.Filter,
        '--artifacts', $artifactsPath
    )
    if ($RunMode -eq 'quick') {
        $arguments += @('--job', 'Dry')
    }
    if ($AffinityMask -ne 0) {
        $arguments += @('--affinity', $AffinityMask.ToString([Globalization.CultureInfo]::InvariantCulture))
    }

    Push-Location -LiteralPath $repositoryRoot
    $previousPdfCorpusRoot = $env:OFFICEIMO_PDF_CORPUS_ROOT
    try {
        if ($name -eq 'pdfcorpusread') {
            $env:OFFICEIMO_PDF_CORPUS_ROOT = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PdfCorpusRoot)
        }
        & dotnet @arguments
        $benchmarkExitCode = $LASTEXITCODE
    } finally {
        $env:OFFICEIMO_PDF_CORPUS_ROOT = $previousPdfCorpusRoot
        Pop-Location
    }
    if ($benchmarkExitCode -ne 0) {
        throw "$name benchmark run failed with exit code $benchmarkExitCode."
    }

    $sizeEvidencePath = $null
    if ($definition.ValidatedSizeEvidence) {
        $sizeEvidencePath = Join-Path $artifactsPath 'validated-size-evidence.json'
        $sizeArguments = @(
            'run',
            '-c', 'Release',
            '-f', $Framework,
            '--project', (Join-Path $repositoryRoot $definition.Project),
            '--',
            'validate',
            '--json', $sizeEvidencePath
        )
        Push-Location -LiteralPath $repositoryRoot
        try {
            & dotnet @sizeArguments
            $sizeEvidenceExitCode = $LASTEXITCODE
        } finally {
            Pop-Location
        }
        if ($sizeEvidenceExitCode -ne 0) {
            throw "$name validated size evidence failed with exit code $sizeEvidenceExitCode."
        }

        $sizeResults = Get-Content -LiteralPath $sizeEvidencePath -Raw | ConvertFrom-Json
        $sizeEvidence = [ordered]@{
            Workload = $name
            ComparisonId = $definition.ComparisonId
            SourceCommit = $gitSha
            Framework = $Framework
            Platform = $platform
            CapturedAtUtc = [DateTimeOffset]::UtcNow.ToString('O')
            Results = $sizeResults
        }
        $sizeEvidenceJson = $sizeEvidence | ConvertTo-Json -Depth 12
        [System.IO.File]::WriteAllText(
            $sizeEvidencePath,
            $sizeEvidenceJson,
            [System.Text.UTF8Encoding]::new($false))
    }
    $provenanceCapture |
        Complete-BenchmarkProvenanceCapture |
        Out-Null

    $result = Import-BenchmarkResult -Path $artifactsPath -Suite $definition.Suite
    $successfulCases = @(
        $result.Summary |
            Where-Object {
                $_.Status -in @('Success', 'Succeeded') -and
                $_.SampleCount -gt 0 -and
                $null -ne $_.MedianMs
            } |
            ForEach-Object {
                Get-BenchmarkEvidenceCaseIdentity `
                    -Row $_ `
                    -VariableName $definition.IdentityVariables
            } |
            Sort-Object -Unique
    )
    $missingCases = @(
        $definition.ExpectedCases |
            Where-Object { $_ -notin $successfulCases }
    )
    $unexpectedCases = @(
        $successfulCases |
            Where-Object { $_ -notin $definition.ExpectedCases }
    )
    if ($missingCases.Count -gt 0 -or $unexpectedCases.Count -gt 0) {
        $details = @()
        if ($missingCases.Count -gt 0) {
            $details += "missing successful cases: $($missingCases -join ', ')"
        }
        if ($unexpectedCases.Count -gt 0) {
            $details += "unexpected cases: $($unexpectedCases -join ', ')"
        }
        throw "$name benchmark evidence is incomplete ($($details -join '; '))."
    }

    $evidenceLocation = Get-BenchmarkEvidenceLocation `
        -ComparisonId $definition.ComparisonId `
        -Platform $platform `
        -RunMode $RunMode `
        -StaticRoot $staticRoot
    $normalizedPath = Join-Path $artifactsPath 'normalized-result.json'
    Write-BenchmarkEvidenceResult -Path $normalizedPath -InputObject $result

    $measurements.Add([pscustomobject]@{
        Workload = $name
        Definition = $definition
        Result = $result
        EvidenceLocation = $evidenceLocation
        ArtifactsPath = $artifactsPath
        NormalizedResult = $normalizedPath
        SizeEvidence = $sizeEvidencePath
        CatalogEligible = $workloadPlan.WillCatalog
    })
}

if (($measurements.Count -ne $selected.Count) -or
    (@($measurements | Where-Object { $null -eq $_.Result }).Count -gt 0)) {
    throw 'Benchmark measurement collection did not produce exactly one normalized result per selected workload.'
}

if ($catalogEligible) {
    foreach ($measurement in @($measurements | Where-Object CatalogEligible)) {
        Update-BenchmarkEvidenceCatalog `
            -InputObject $measurement.Result `
            -Path $catalogPath `
            -ComparisonId $measurement.Definition.ComparisonId `
            -ResultPath $measurement.EvidenceLocation.ResultPath `
            -ResultArtifactPath $measurement.EvidenceLocation.Path `
            -RunMode $RunMode `
            -Platform $platform `
            -ExpectedPlatform windows, linux, macos `
            -Publish:$Publish | Out-Null
    }
}

$outputs = foreach ($measurement in $measurements) {
    [pscustomobject]@{
        Workload = $measurement.Workload
        Platform = $platform
        RunMode = $RunMode
        Publish = $executionPlanByWorkload[$measurement.Workload].Publish
        AffinityMask = $affinityLabel
        SourceCommit = $gitSha
        ArtifactsPath = $measurement.ArtifactsPath
        NormalizedResult = if ($measurement.CatalogEligible) {
            $measurement.EvidenceLocation.Path
        } else {
            $measurement.NormalizedResult
        }
        SizeEvidence = $measurement.SizeEvidence
        EvidenceCatalog = if ($measurement.CatalogEligible) { $catalogPath } else { $null }
    }
}

$outputs
