[CmdletBinding()]
param(
    [string] $Framework = 'net10.0',
    [switch] $SkipGeneration
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$examplesProject = Join-Path $repoRoot 'OfficeIMO.Examples\OfficeIMO.Examples.csproj'
$readerProject = Join-Path $repoRoot 'OfficeIMO.Reader.Tool\OfficeIMO.Reader.Tool.csproj'
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
    $readerAssembly = Join-Path $repoRoot "OfficeIMO.Reader.Tool\bin\Debug\$Framework\OfficeIMO.Reader.Tool.dll"
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = 'dotnet'
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in @(
        $readerAssembly, 'read', '-', '--name', 'design-brief.pptx',
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
            throw "OfficeIMO.Reader.Tool failed: $($process.StandardError.ReadToEnd())"
        }
    } finally {
        $process.Dispose()
    }
}

if (-not $SkipGeneration) {
    foreach ($exampleSwitch in @(
        '--powerpoint-design-brief',
        '--pdf-showcase',
        '--excel-report-workflow',
        '--onenote',
        '--visio-premium'
    )) {
        Invoke-DotNet @(
            'run', '--project', $examplesProject, '-f', $Framework, '--', $exampleSwitch
        )
    }
}

$powerPointPath = Join-Path $documentsRoot 'PowerPoint Design Brief Recommendations.pptx'
$readerPath = Join-Path $documentsRoot 'PowerPoint-Design-Brief.reader.public.json'
New-ReaderProjection -InputPath $powerPointPath -OutputPath $readerPath

$artifacts = @(
    [ordered]@{
        id = 'powerpoint-output'
        source = $powerPointPath
        destination = 'powerpoint/design-brief-recommendations.pptx'
        generator = 'dotnet run --project OfficeIMO.Examples -f net10.0 -- --powerpoint-design-brief'
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
        generator = 'dotnet run --project OfficeIMO.Examples -f net10.0 -- --pdf-showcase'
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
        id = 'excel-output'
        source = (Join-Path $documentsRoot 'ExcelReportWorkflow.xlsx')
        destination = 'excel/report-workflow.xlsx'
        generator = 'dotnet run --project OfficeIMO.Examples -f net10.0 -- --excel-report-workflow'
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
        generator = 'dotnet run --project OfficeIMO.Examples -f net10.0 -- --visio-premium'
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
        generator = 'OfficeIMO.Reader.Tool read - --name design-brief.pptx --format json'
        evidence = 'Schema-versioned Reader result generated from the downloadable PPTX.'
    },
    [ordered]@{
        id = 'onenote-section'
        source = (Join-Path $documentsRoot 'OfficeIMO-OneNote.one')
        destination = 'onenote/offline-planning.one'
        generator = 'dotnet run --project OfficeIMO.Examples -f net10.0 -- --onenote'
        evidence = 'Native offline OneNote section.'
    },
    [ordered]@{
        id = 'onenote-package'
        source = (Join-Path $documentsRoot 'OfficeIMO-OneNote.onepkg')
        destination = 'onenote/offline-planning.onepkg'
        generator = 'dotnet run --project OfficeIMO.Examples -f net10.0 -- --onenote'
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
    Copy-Evidence -Source $artifact.source -Destination $destination
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
