$family = Get-BenchmarkInput Family -Required
$cells = Get-BenchmarkInput Cells 12,120,1200 -Int
$codeCharacters = Get-BenchmarkInput CodeLineCharacters 32 -Int
$outputCharacters = Get-BenchmarkInput OutputCharacters 32 -Int
$imagePixels = Get-BenchmarkInput RichImagePixels 64 -Int
$timeout = Get-BenchmarkInput WorkerTimeoutSeconds 60 -Int
$session = Get-BenchmarkInput Session -Required
$planning = Get-BenchmarkInput Planning $false -Bool
$cold = Get-BenchmarkInput Cold $false -Bool
$binary = Get-BenchmarkInput ManagedBinary -Required
$python = Get-BenchmarkInput PythonPath ''
$dotnet = Get-BenchmarkInput DotNetPath ''
$pdfText = Get-BenchmarkInput PdfText ''
$pdfImages = Get-BenchmarkInput PdfImages ''
$inputsRoot = Get-BenchmarkInput InputsRoot -Required
$fixtureRoot = Join-Path $PSScriptRoot '../../OfficeIMO.TestAssets/Notebook/Jupyter'
$referenceScript = Join-Path $PSScriptRoot 'notebook-reference.py'
$verifyScript = Join-Path $PSScriptRoot 'notebook-verify.py'
$manifest = Get-Content -LiteralPath (Join-Path $fixtureRoot 'fixtures.json') -Raw | ConvertFrom-Json
$availability = @{
    OfficeIMO = Get-BenchmarkInput HasOfficeIMO $false -Bool
    Microsoft = Get-BenchmarkInput HasMicrosoft $false -Bool
    Verso = Get-BenchmarkInput HasVerso $false -Bool
    Reference = Get-BenchmarkInput HasNbformat $false -Bool
}
$hasNbconvert = Get-BenchmarkInput HasNbconvert $false -Bool
$hasChromium = Get-BenchmarkInput HasChromium $false -Bool

New-BenchmarkSuite "officeimo-notebook-$($family.ToLowerInvariant())" {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 5 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Family $family
    if ($cold) {
        Add-BenchmarkMetadata Measurement 'PowerForge includes fresh worker process startup, imports, initial model preparation, and the first requested operation. Export also writes the output file and associated assets during timing. OS caches are not cleared. Operation-only managed allocation excludes startup, preparation and destination writes. Input generation and artifact readback remain outside measurement.'
    } else {
        Add-BenchmarkMetadata Measurement 'PowerForge elapsed time includes warmed worker IPC. Input generation, preloaded export input, and validation are outside measurement. Worker managed allocation, when present, is process-wide .NET allocation per call. PowerForge host allocation is harness overhead and does not measure the workers.'
    }
    Add-BenchmarkMetadata WorkerMemory 'Working set and sampled memory from PowerForge observe its own host. Managed worker allocation covers the measured owner call only. Peak RSS for worker processes and Chromium descendants is unqualified.'
    Add-BenchmarkMetadata CodecContract 'Parse opens a public notebook model from the exact bytes and verifies independently recovered contents after timing. Each engine can defer projection, so parse cost is not an eager materialization comparison. RoundTrip requires the same on-disk semantics including metadata, MIME alternatives and IDs.'
    Add-BenchmarkMetadata ProjectionContract 'Open exact bytes and materialize ordered cell types and complete source strings, excluding output projection. All engines must recover identical cells. Microsoft and Verso decode the given UTF-8 bytes for their string APIs inside the measured operation.'
    Add-BenchmarkMetadata ExportContract 'Markdown and HTML start from a prepared notebook model. PDF includes notebook parsing, static conversion, and complete PDF bytes. The Python WebPDF path uses the lab template on synthetic inputs, disabled kernel execution and an offline security policy. Compare artifact witnesses, and inspect visual fidelity separately.'
    Add-BenchmarkMetadata FidelityContract 'Fidelity is a compatibility observation, with one probe per selected engine and input. Inspect each fidelity.json. Discard timings from this family for comparative claims.'
    Add-BenchmarkMetadata MicrosoftVersion '1.0.0-beta.26120.1 (MIT, df4e5cbcbe1b5d0ea85247226b7b896765a54611)'
    Add-BenchmarkMetadata VersoVersion '1.2.3 (MIT, 6f2e3dd9103be27931340bcb0bba2e326590f992)'
    Add-BenchmarkMetadata ManagedRuntime (Get-BenchmarkInput ManagedRuntime 'Unavailable or not selected')
    Add-BenchmarkMetadata PythonEnvironment (Get-BenchmarkInput PythonEnvironment 'Unavailable in plan mode')
    Add-BenchmarkMetadata FixtureRevision $manifest.revision
    Add-BenchmarkMetadata SpecSha256 (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'notebook.benchmark.ps1') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WrapperSha256 (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Run-NotebookBenchmarks.ps1') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WorkerTransportSha256 (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'notebook-worker.psm1') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata RunnerSha256 (Get-FileHash (Get-Command Invoke-BenchmarkSuite).ImplementingType.Assembly.Location -Algorithm SHA256).Hash
    Add-BenchmarkMetadata ReferenceSha256 (Get-FileHash -LiteralPath $referenceScript -Algorithm SHA256).Hash
    Add-BenchmarkMetadata ValidatorSha256 (Get-FileHash -LiteralPath $verifyScript -Algorithm SHA256).Hash
    if ($availability.OfficeIMO) {
        Add-BenchmarkMetadata WorkloadSha256 (Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash
        $assemblies = @('OfficeIMO.Notebook', 'OfficeIMO.Notebook.Html', 'OfficeIMO.Notebook.Markdown', 'OfficeIMO.Notebook.Pdf')
        if ($availability.Microsoft) { $assemblies += 'Microsoft.DotNet.Interactive.Documents' }
        if ($availability.Verso) { $assemblies += 'Verso', 'Verso.Abstractions' }
        foreach ($name in $assemblies) {
            Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash -LiteralPath (Join-Path (Split-Path $binary) "$name.dll") -Algorithm SHA256).Hash
        }
    }
    if ($pdfText -and $pdfImages) {
        Add-BenchmarkMetadata PopplerTextSha256 (Get-FileHash -LiteralPath $pdfText -Algorithm SHA256).Hash
        Add-BenchmarkMetadata PopplerImagesSha256 (Get-FileHash -LiteralPath $pdfImages -Algorithm SHA256).Hash
    }
    Add-BenchmarkCaseSource {
        if ($family -eq 'Fidelity') {
            Assert-BenchmarkValue -Actual (Get-FileHash -LiteralPath (Join-Path $fixtureRoot 'LICENSE') -Algorithm SHA256).Hash.ToLowerInvariant() -Expected $manifest.licenseSha256
            foreach ($fixture in $manifest.notebooks) {
                $path = Join-Path $fixtureRoot $fixture.file
                Assert-BenchmarkValue -Actual (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -Expected $fixture.sha256
                [pscustomobject]@{ Name = $fixture.name; InputFile = $fixture.file; InputSha256 = $fixture.sha256; Corpus = 'Jupyter upstream' }
            }
            return
        }
        foreach ($size in $cells) {
            $shapes = if ($family -in 'Codec', 'Export') { 'Common', 'Rich' } else { 'Common' }
            foreach ($shape in $shapes) {
                $pixels = if ($shape -eq 'Rich') { $imagePixels } else { 0 }
                $name = "$shape-$size"
                $file = "$name-$codeCharacters-$outputCharacters-$pixels.ipynb"
                $path = Join-Path $inputsRoot $file
                $hash = if ($planning) { 'Generated on execution' } else {
                    & $python $referenceScript generate $path $size $codeCharacters $outputCharacters $pixels
                }
                [pscustomobject]@{
                    Name = $name; InputFile = $file; InputSha256 = $hash; Corpus = 'Synthetic'; CellCount = $size
                    CodeLineCharacters = $codeCharacters; OutputCharacters = $outputCharacters; ImagePixels = $pixels
                }
            }
        }
    }
    Add-BenchmarkSkipRule {
        param($case)
        if ($case.Engine -in 'Microsoft', 'Verso' -and $case.Scenario -eq 'Jupyter3') { return $true }
        if ($planning) { return }
        if (-not $availability[[string] $case.Engine]) { return $true }
        if ($case.Operation -eq 'Pdf' -and (-not $pdfText -or -not $pdfImages)) { return $true }
        if ($case.Engine -eq 'Reference' -and $case.Operation -in 'Markdown', 'Html', 'Pdf' -and -not $hasNbconvert) { return $true }
        if ($case.Engine -eq 'Reference' -and $case.Operation -eq 'Pdf' -and -not $hasChromium) { return $true }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Executable = if ($case.Engine -eq 'Reference') { $python } else { $dotnet }
        $run.Arguments = if ($case.Engine -eq 'Reference') { @($referenceScript, 'worker') } else { @($binary, 'worker') }
        $run.InputPath = if ($family -eq 'Fidelity') {
            Join-Path $fixtureRoot $case.InputFile
        } else {
            Join-Path $inputsRoot $case.InputFile
        }
        $run.WorkerName = if ($cold) { [Guid]::NewGuid().ToString('N') } else { [string]$case.Engine }
        $extension = switch ($case.Operation) {
            Html { 'html' }
            Markdown { 'md' }
            Pdf { 'pdf' }
            default { 'json' }
        }
        $run.Artifact = Join-Path $run.OutputDirectory "output.$extension"
        $run.PrepareRequest = @{ action = 'prepare'; engine = [string]$case.Engine; operation = [string]$case.Operation
            input = $run.InputPath; artifact = $run.Artifact }
        if (-not $cold) {
            $null = Invoke-NotebookWorker $session $run.WorkerName $run.Executable $run.Arguments $run.PrepareRequest $timeout
        }
    }
    $operations = @{
        OfficeIMO = @('Parse', 'RoundTrip', 'SourceProjection', 'RichRoundTrip', 'Html', 'Markdown', 'Pdf')
        Reference = @('Parse', 'RoundTrip', 'SourceProjection', 'RichRoundTrip', 'Html', 'Markdown', 'Pdf')
        Microsoft = @('SourceProjection', 'RichRoundTrip')
        Verso = @('SourceProjection', 'RichRoundTrip')
    }
    foreach ($engine in $operations.Keys) {
        Add-BenchmarkEngine $engine {
            foreach ($operation in $operations[$engine]) {
                Add-BenchmarkOperation $operation {
                    param($case, $run)
                    if ($cold) {
                        $null = Invoke-NotebookWorker $session $run.WorkerName $run.Executable $run.Arguments $run.PrepareRequest $timeout
                    }
                    $null = Invoke-NotebookWorker $session $run.WorkerName $run.Executable $run.Arguments @{ action = 'run' } $timeout
                    if ($cold -and $family -eq 'Export') {
                        $run.Worker = Invoke-NotebookWorker $session $run.WorkerName $run.Executable $run.Arguments @{ action = 'finish' } $timeout
                    }
                }
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        try {
            if (-not $run.Worker) {
                $run.Worker = Invoke-NotebookWorker $session $run.WorkerName $run.Executable $run.Arguments @{ action = 'finish' } $timeout
            }
        } finally {
            if ($cold) { Stop-NotebookWorkers -Session $session -Name $run.WorkerName }
        }
        $verifyArguments = @($verifyScript, $case.Operation, $run.InputPath, $run.Artifact, "--text-tool=$pdfText", "--image-tool=$pdfImages")
        $run.Verified = & $python @verifyArguments | ConvertFrom-Json
        Assert-BenchmarkValue -Actual $run.Verified.inputSha256 -Expected $case.InputSha256
        Assert-BenchmarkValue -Actual $run.Verified.outputBytes -Expected $run.Worker.outputBytes
    }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Verified.inputBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.Verified.outputBytes }
    Add-BenchmarkMetric VerifiedCells { param($case, $run) $run.Verified.verifiedCells }
    Add-BenchmarkMetric VerifiedPages { param($case, $run) $run.Verified.verifiedPages }
    Add-BenchmarkMetric VerifiedImages { param($case, $run) $run.Verified.verifiedImages }
    Add-BenchmarkMetric SemanticDifferences { param($case, $run) $run.Verified.semanticDifferences }
    Add-BenchmarkMetric WorkerManagedAllocatedBytes {
        param($case, $run)
        if ($null -ne $run.Worker.allocatedBytes) { $run.Worker.allocatedBytes }
    }
    if ($family -in 'Projection', 'Export') {
        Add-BenchmarkComparison Engine -Baseline OfficeIMO -Metric MedianMs
    }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
