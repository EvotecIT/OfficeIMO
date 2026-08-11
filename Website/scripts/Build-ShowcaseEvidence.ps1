[CmdletBinding()]
param(
    [string] $Framework = 'net10.0',
    [switch] $SkipGeneration,
    [switch] $ManifestOnly
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$examplesProject = Join-Path $repoRoot 'OfficeIMO.Examples\OfficeIMO.Examples.csproj'
$readerProject = Join-Path $repoRoot 'OfficeIMO.Tool\OfficeIMO.Tool.csproj'
$documentsRoot = Join-Path $repoRoot "OfficeIMO.Examples\bin\Debug\$Framework\Documents"
$downloadRoot = Join-Path $repoRoot 'Website\static\downloads\showcase'
$manifestPath = Join-Path $downloadRoot 'manifest.json'

function Invoke-DotNet {
    param([Parameter(Mandatory)][string[]] $Arguments)

    & dotnet @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet failed with exit code $LASTEXITCODE."
    }
}

function Copy-Evidence {
    param(
        [Parameter(Mandatory)][string] $Source,
        [Parameter(Mandatory)][string] $Destination
    )

    if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) {
        throw "Expected showcase evidence was not generated: $Source"
    }

    $destinationDirectory = Split-Path -Parent $Destination
    New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
    Copy-Item -LiteralPath $Source -Destination $Destination -Force
}

function New-ReaderProjection {
    param(
        [Parameter(Mandatory)][string] $InputPath,
        [Parameter(Mandatory)][string] $OutputPath
    )

    Invoke-DotNet @('build', $readerProject, '-f', $Framework, '--nologo')
    $readerAssembly = Join-Path $repoRoot "OfficeIMO.Tool\bin\Debug\$Framework\OfficeIMO.Tool.dll"
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = 'dotnet'
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in @(
        $readerAssembly, 'reader', 'read', '-', '--name', 'design-brief.pptx',
        '--format', 'json', '--output', $OutputPath
    )) {
        [void] $startInfo.ArgumentList.Add($argument)
    }

    $process = [System.Diagnostics.Process]::Start($startInfo)
    try {
        $inputBytes = [System.IO.File]::ReadAllBytes($InputPath)
        $process.StandardInput.BaseStream.Write($inputBytes, 0, $inputBytes.Length)
        $process.StandardInput.Close()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) {
            throw "OfficeIMO.Tool reader command failed: $($process.StandardError.ReadToEnd())"
        }
    } finally {
        $process.Dispose()
    }
}

function New-PdfPagePreview {
    param(
        [Parameter(Mandatory)][string] $InputPath,
        [Parameter(Mandatory)][string] $OutputPath
    )

    $renderer = Get-Command 'pdftocairo.exe' -ErrorAction SilentlyContinue
    if ($null -eq $renderer) {
        $renderer = Get-Command 'pdftocairo' -ErrorAction SilentlyContinue
    }
    if ($null -eq $renderer) {
        throw 'pdftocairo is required to refresh independent PDF page previews. Install Poppler or use -ManifestOnly with committed evidence.'
    }
    if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
        throw "Expected showcase PDF was not generated: $InputPath"
    }

    $outputDirectory = Split-Path -Parent $OutputPath
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    $outputBase = Join-Path $outputDirectory ([System.IO.Path]::GetFileNameWithoutExtension($OutputPath))
    & $renderer.Source '-f' '1' '-l' '1' '-singlefile' '-png' '-r' '120' $InputPath $outputBase
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
        throw "pdftocairo failed to render '$InputPath'."
    }
}

if (-not $SkipGeneration -and -not $ManifestOnly) {
    foreach ($exampleSwitch in @(
        '--powerpoint-design-brief',
        '--pdf-showcase',
        '--html-invoice',
        '--html-renderer-gallery',
        '--excel-report-workflow',
        '--onenote',
        '--visio-premium',
        '--showcase-real-world'
    )) {
        Invoke-DotNet @(
            'run', '--project', $examplesProject, '-f', $Framework, '--', $exampleSwitch
        )
    }
}

$powerPointPath = Join-Path $documentsRoot 'PowerPoint Design Brief Recommendations.pptx'
$readerPath = Join-Path $documentsRoot 'PowerPoint-Design-Brief.reader.public.json'
if (-not $ManifestOnly) {
    New-ReaderProjection -InputPath $powerPointPath -OutputPath $readerPath
}
$workflowRoot = Join-Path $documentsRoot 'WorkflowShowcase'
$realWorldWorkflows = @(
    'customer-delivery-summary',
    'change-approval-memo',
    'release-readiness-report',
    'project-handover-brief'
)
if (-not $ManifestOnly) {
    foreach ($workflow in $realWorldWorkflows) {
        New-PdfPagePreview `
            -InputPath (Join-Path $workflowRoot "$workflow.pdf") `
            -OutputPath (Join-Path $workflowRoot "$workflow.png")
    }
}

$artifacts = @(
    [ordered]@{
        id = 'powerpoint-output'
        source = $powerPointPath
        destination = 'powerpoint/design-brief-recommendations.pptx'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --powerpoint-design-brief"
        evidence = 'Editable PPTX; Open XML validation is part of the example.'
    },
    [ordered]@{
        id = 'powerpoint-preview'
        source = (Join-Path $repoRoot 'Website\static\images\powerpoint\examples\design-brief-selected.png')
        destination = 'powerpoint/design-brief-selected.png'
        generator = 'OfficeIMO.PowerPoint design-brief rendering proof'
        evidence = 'Rendered selected-direction slide from the same example.'
    },
    [ordered]@{
        id = 'pdf-output'
        source = (Join-Path $documentsRoot 'Pdf.Showcase.Dashboard.pdf')
        destination = 'pdf/showcase-dashboard.pdf'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --pdf-showcase"
        evidence = 'First-party PDF dashboard output.'
    },
    [ordered]@{
        id = 'pdf-preview'
        source = (Join-Path $repoRoot 'OfficeIMO.Pdf.Tests\Pdf\VisualBaselines\officeimo-pdf-showcase-dashboard.page1.poppler.png')
        destination = 'pdf/showcase-dashboard-page1.png'
        generator = 'OfficeIMO.Pdf visual baseline rendered with Poppler'
        evidence = 'Approved page-one visual baseline for the generated dashboard.'
    },
    [ordered]@{
        id = 'html-pdf-output'
        source = (Join-Path $documentsRoot 'HtmlInvoiceShowcase.pdf')
        destination = 'html/invoice.pdf'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --html-invoice"
        evidence = 'Three-page searchable purchase report with wrapped rows, repeated table headers, page counters, and one totals block.'
    },
    [ordered]@{
        id = 'html-png-output'
        source = (Join-Path $documentsRoot 'HtmlInvoiceShowcase.png')
        destination = 'html/invoice.png'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --html-invoice"
        evidence = 'First-page PNG generated from the same parsed HTML and options object as the multi-page PDF.'
    },
    [ordered]@{
        id = 'html-svg-output'
        source = (Join-Path $documentsRoot 'HtmlInvoiceShowcase.svg')
        destination = 'html/invoice.svg'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --html-invoice"
        evidence = 'First-page SVG generated from the same parsed HTML and options object as the multi-page PDF.'
    },
    [ordered]@{
        id = 'html-renderer-gallery-source'
        source = (Join-Path $documentsRoot 'HtmlManagedRendererGallery.html')
        destination = 'html/managed-renderer-gallery.html'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --html-renderer-gallery"
        evidence = 'Authored HTML and CSS source used for every managed renderer gallery output.'
    },
    [ordered]@{
        id = 'html-renderer-gallery-pdf'
        source = (Join-Path $documentsRoot 'HtmlManagedRendererGallery.pdf')
        destination = 'html/managed-renderer-gallery.pdf'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --html-renderer-gallery"
        evidence = 'Searchable managed PDF with native AcroForm controls and page-margin content.'
    },
    [ordered]@{
        id = 'html-renderer-gallery-png'
        source = (Join-Path $documentsRoot 'HtmlManagedRendererGallery.png')
        destination = 'html/managed-renderer-gallery.png'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --html-renderer-gallery"
        evidence = 'Deterministic first-page PNG generated from the same managed scene as the PDF.'
    },
    [ordered]@{
        id = 'html-renderer-gallery-svg'
        source = (Join-Path $documentsRoot 'HtmlManagedRendererGallery.svg')
        destination = 'html/managed-renderer-gallery.svg'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --html-renderer-gallery"
        evidence = 'Vector SVG generated from the shared backend-neutral scene.'
    },
    [ordered]@{
        id = 'word-workflow-source'
        source = (Join-Path $workflowRoot 'customer-delivery-summary.docx')
        destination = 'word/customer-delivery-summary.docx'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --showcase-real-world"
        evidence = 'Editable DOCX generated from the same Word model as the PDF conversion.'
    },
    [ordered]@{
        id = 'word-workflow-pdf'
        source = (Join-Path $workflowRoot 'customer-delivery-summary.pdf')
        destination = 'word/customer-delivery-summary.pdf'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --showcase-real-world"
        evidence = 'First-party Word-to-PDF conversion of the downloadable DOCX.'
    },
    [ordered]@{
        id = 'word-workflow-preview'
        source = (Join-Path $workflowRoot 'customer-delivery-summary.png')
        destination = 'word/customer-delivery-summary-page1.png'
        generator = 'Poppler pdftocairo page-one render of the generated Word PDF'
        evidence = 'Independent page-one render of the first-party Word-to-PDF result.'
    },
    [ordered]@{
        id = 'rtf-workflow-source'
        source = (Join-Path $workflowRoot 'change-approval-memo.rtf')
        destination = 'rtf/change-approval-memo.rtf'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --showcase-real-world"
        evidence = 'RTF memo generated from the same model as the PDF conversion.'
    },
    [ordered]@{
        id = 'rtf-workflow-pdf'
        source = (Join-Path $workflowRoot 'change-approval-memo.pdf')
        destination = 'rtf/change-approval-memo.pdf'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --showcase-real-world"
        evidence = 'First-party RTF-to-PDF conversion of the downloadable memo.'
    },
    [ordered]@{
        id = 'rtf-workflow-preview'
        source = (Join-Path $workflowRoot 'change-approval-memo.png')
        destination = 'rtf/change-approval-memo-page1.png'
        generator = 'Poppler pdftocairo page-one render of the generated RTF PDF'
        evidence = 'Independent page-one render of the first-party RTF-to-PDF result.'
    },
    [ordered]@{
        id = 'markdown-workflow-source'
        source = (Join-Path $workflowRoot 'release-readiness-report.md')
        destination = 'markdown/release-readiness-report.md'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --showcase-real-world"
        evidence = 'Portable Markdown source generated from the same document model as the PDF.'
    },
    [ordered]@{
        id = 'markdown-workflow-pdf'
        source = (Join-Path $workflowRoot 'release-readiness-report.pdf')
        destination = 'markdown/release-readiness-report.pdf'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --showcase-real-world"
        evidence = 'Themed first-party Markdown-to-PDF conversion.'
    },
    [ordered]@{
        id = 'markdown-workflow-preview'
        source = (Join-Path $workflowRoot 'release-readiness-report.png')
        destination = 'markdown/release-readiness-report-page1.png'
        generator = 'Poppler pdftocairo page-one render of the generated Markdown PDF'
        evidence = 'Independent page-one render of the themed Markdown PDF.'
    },
    [ordered]@{
        id = 'opendocument-workflow-source'
        source = (Join-Path $workflowRoot 'project-handover-brief.odt')
        destination = 'open-document/project-handover-brief.odt'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --showcase-real-world"
        evidence = 'Vendor-neutral ODT source generated from the OpenDocument model.'
    },
    [ordered]@{
        id = 'opendocument-workflow-pdf'
        source = (Join-Path $workflowRoot 'project-handover-brief.pdf')
        destination = 'open-document/project-handover-brief.pdf'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --showcase-real-world"
        evidence = 'First-party OpenDocument-to-PDF conversion of the downloadable ODT.'
    },
    [ordered]@{
        id = 'opendocument-workflow-preview'
        source = (Join-Path $workflowRoot 'project-handover-brief.png')
        destination = 'open-document/project-handover-brief-page1.png'
        generator = 'Poppler pdftocairo page-one render of the generated OpenDocument PDF'
        evidence = 'Independent page-one render of the first-party OpenDocument-to-PDF result.'
    },
    [ordered]@{
        id = 'excel-output'
        source = (Join-Path $documentsRoot 'ExcelReportWorkflow.xlsx')
        destination = 'excel/report-workflow.xlsx'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --excel-report-workflow"
        evidence = 'Editable XLSX with formulas, chart, table, and pivot metadata.'
    },
    [ordered]@{
        id = 'excel-preview'
        source = (Join-Path $documentsRoot 'ExcelReportWorkflow.png')
        destination = 'excel/report-workflow.png'
        generator = 'OfficeIMO.Excel range ExportImage'
        evidence = 'Dependency-free rendered worksheet range.'
    },
    [ordered]@{
        id = 'excel-preflight'
        source = (Join-Path $documentsRoot 'ExcelReportWorkflow.preflight.txt')
        destination = 'excel/report-workflow.preflight.txt'
        generator = 'OfficeIMO.Excel InspectFeatures'
        evidence = 'Real blocked-PDF diagnostic for unmaterialized pivot output.'
    },
    [ordered]@{
        id = 'visio-output'
        source = (Join-Path $documentsRoot 'Premium Visio Showcase\Premium - Cloud Architecture.vsdx')
        destination = 'visio/premium-cloud-architecture.vsdx'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --visio-premium"
        evidence = 'Editable VSDX validated by the premium gallery example.'
    },
    [ordered]@{
        id = 'visio-preview'
        source = (Join-Path $repoRoot 'OfficeIMO.Visio.Tests\Visio\VisualBaselines\officeimo-visio-premium-cloud-architecture-native-page1.png')
        destination = 'visio/premium-cloud-architecture-page1.png'
        generator = 'OfficeIMO.Visio dependency-free native PNG renderer'
        evidence = 'Approved native-renderer baseline for the same premium scenario.'
    },
    [ordered]@{
        id = 'reader-output'
        source = $readerPath
        destination = 'reader/design-brief.reader.json'
        generator = 'officeimo reader read - --name design-brief.pptx --format json'
        evidence = 'Schema-versioned Reader result generated from the downloadable PPTX.'
    },
    [ordered]@{
        id = 'onenote-section'
        source = (Join-Path $documentsRoot 'OfficeIMO-OneNote.one')
        destination = 'onenote/offline-planning.one'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --onenote"
        evidence = 'Native offline OneNote section.'
    },
    [ordered]@{
        id = 'onenote-package'
        source = (Join-Path $documentsRoot 'OfficeIMO-OneNote.onepkg')
        destination = 'onenote/offline-planning.onepkg'
        generator = "dotnet run --project OfficeIMO.Examples -f $Framework -- --onenote"
        evidence = 'Native offline OneNote notebook package.'
    },
    [ordered]@{
        id = 'onenote-pdf'
        source = (Join-Path $documentsRoot 'OfficeIMO-OneNote.pdf')
        destination = 'onenote/offline-planning.pdf'
        generator = 'OfficeIMO.OneNote.Pdf SaveAsPdf'
        evidence = 'PDF export of the same generated section.'
    },
    [ordered]@{
        id = 'onenote-html'
        source = (Join-Path $documentsRoot 'OfficeIMO-OneNote.html')
        destination = 'onenote/offline-planning.html.txt'
        generator = 'OfficeIMO.OneNote.Html SaveAsHtml'
        evidence = 'HTML export used by the code-native gallery preview.'
    },
    [ordered]@{
        id = 'onenote-markdown'
        source = (Join-Path $documentsRoot 'OfficeIMO-OneNote.md')
        destination = 'onenote/offline-planning.md'
        generator = 'OfficeIMO.OneNote.Markdown ToMarkdown'
        evidence = 'Markdown export of the same generated section.'
    }
)

$manifestArtifacts = foreach ($artifact in $artifacts) {
    $destination = Join-Path $downloadRoot $artifact.destination
    if ($ManifestOnly) {
        if (-not (Test-Path -LiteralPath $destination -PathType Leaf)) {
            throw "Expected showcase evidence is missing: $destination"
        }
    } else {
        Copy-Evidence -Source $artifact.source -Destination $destination
    }
    $file = Get-Item -LiteralPath $destination
    [ordered]@{
        id = $artifact.id
        path = '/downloads/showcase/' + ($artifact.destination -replace '\\', '/')
        bytes = $file.Length
        sha256 = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant()
        generator = $artifact.generator
        evidence = $artifact.evidence
    }
}

$manifestDirectory = Split-Path -Parent $manifestPath
New-Item -ItemType Directory -Path $manifestDirectory -Force | Out-Null
[ordered]@{
    schema = 'officeimo.showcase-evidence'
    schemaVersion = 1
    artifacts = @($manifestArtifacts)
} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding utf8NoBOM

Write-Host "Showcase evidence refreshed: $($manifestArtifacts.Count) artifacts"
Write-Host "Manifest: $manifestPath"
