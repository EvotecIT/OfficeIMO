$mode = Get-BenchmarkInput Mode -Required
$cellCounts = Get-BenchmarkInput Cells 12, 120, 1200 -Int
$filesPerCase = Get-BenchmarkInput FilesPerCase 1 -Int
$throttleLimit = Get-BenchmarkInput ThrottleLimit 4 -Int
$fixtureRoot = Get-BenchmarkInput FixtureRoot -Required
$moduleManifest = Get-BenchmarkInput ModuleManifest -Required
$peerManifest = Get-BenchmarkInput PowerShellNotebookManifest ''
$powerShellPath = Get-BenchmarkInput PowerShellPath -Required
$observeMemory = Get-BenchmarkInput ObserveMemory $false -Bool
$childScript = Join-Path $PSScriptRoot 'Invoke-NotebookBenchmarkChild.ps1'
$fixtureHelpers = Join-Path $PSScriptRoot 'notebook-performance.fixtures.ps1'
$darwinTime = $observeMemory -and
    [Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([Runtime.InteropServices.OSPlatform]::OSX) -and
    (Test-Path -LiteralPath '/usr/bin/time' -PathType Leaf)

. $fixtureHelpers

New-BenchmarkSuite 'pswriteoffice-notebook' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 5 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Deterministic v4.5 Python notebooks with ordered prose, one Markdown table per prose cell, code and saved streams. No cells execute.'
    Add-BenchmarkMetadata Timing 'Warm measures loaded commands. Cold measures a fresh shell, module import, first Markdown conversion, and process exit. Batch measures PSWriteOffice Markdown conversion at the declared concurrency; fixture generation and content validation are outside timing.'
    Add-BenchmarkMetadata Memory 'PowerForge allocation and sampled memory describe the supervising PowerShell process. Cold worker peak RSS is captured separately in an instrumented run via Darwin time -l or Process.PeakWorkingSet64 where the host supplies it.'
    Add-BenchmarkMetadata PSWriteOfficeManifestSha256 (Get-FileHash -LiteralPath $moduleManifest).Hash
    Add-BenchmarkMetadata FixtureGeneratorSha256 (Get-FileHash -LiteralPath $fixtureHelpers).Hash
    if ($peerManifest) {
        Add-BenchmarkMetadata PowerShellNotebook 'PowerShell Gallery 3.0.0, MIT'
        Add-BenchmarkMetadata PowerShellNotebookManifestSha256 (Get-FileHash -LiteralPath $peerManifest).Hash
        $peerScript = Join-Path (Split-Path $peerManifest) 'PowerShellNotebook.psm1'
        Add-BenchmarkMetadata PowerShellNotebookScriptSha256 (Get-FileHash -LiteralPath $peerScript).Hash
    }
    $converter = Get-Command PSWriteOffice\Convert-OfficeDocument -ListImported -ErrorAction SilentlyContinue
    if ($mode -ne 'Cold' -and $converter) {
        $assembly = $converter.ImplementingType.Assembly.Location
        Add-BenchmarkMetadata PSWriteOfficeAssemblySha256 (Get-FileHash -LiteralPath $assembly).Hash
    }
    Add-BenchmarkCaseSource {
        foreach ($size in $cellCounts) {
            $directory = Join-Path $fixtureRoot "Notebook$size"
            $path = Join-Path $directory 'notebook-1.ipynb'
            $sha = if (Test-Path -LiteralPath $path) { (Get-FileHash -LiteralPath $path).Hash } else { 'planned' }
            [pscustomobject]@{
                Name = "Notebook$size"
                Cells = $size
                Files = $filesPerCase
                ThrottleLimit = if ($mode -eq 'Batch') { $throttleLimit } else { 1 }
                InputPath = $path
                InputSha256 = $sha
            }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)

        if ($case.Operation -eq 'ReadCells') {
            return
        }
        if ($case.Operation -eq 'BatchMarkdown') {
            $run.BatchOutput = Join-Path $run.OutputDirectory 'batch'
            if (Test-Path -LiteralPath $run.BatchOutput) {
                Remove-Item -LiteralPath $run.BatchOutput -Recurse
            }
            [void][IO.Directory]::CreateDirectory($run.BatchOutput)
            $run.BatchInputs = @(for ($index = 1; $index -le $case.Files; $index++) {
                Join-Path (Split-Path $case.InputPath) "notebook-$index.ipynb"
            })
            return
        }
        $extension = if ($case.Operation -eq 'Html') { 'html' } else { 'md' }
        $run.OutputPath = Join-Path $run.OutputDirectory "result.$extension"
        if (Test-Path -LiteralPath $run.OutputPath) {
            Remove-Item -LiteralPath $run.OutputPath
        }
        if ($mode -eq 'Cold') {
            $path = if ($case.Engine -eq 'PSWriteOffice') { $moduleManifest } else { $peerManifest }
            $run.ChildArguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', $childScript,
                '-Engine', $case.Engine, '-ModuleManifest', $path, '-InputPath', $case.InputPath, '-OutputPath', $run.OutputPath,
                '-ObserveMemory', [string]$observeMemory)
            $run.TimeLog = Join-Path $run.OutputDirectory 'worker-time.txt'
        }
    }
    $startup = {
        param($case, $run)

        $arguments = $run.ChildArguments
        if ($darwinTime) {
            $run.ChildOutput = & /usr/bin/time -l $powerShellPath @arguments 2> $run.TimeLog
        } else {
            $run.ChildOutput = & $powerShellPath @arguments
        }
        if ($LASTEXITCODE -ne 0) {
            throw "Notebook worker exited with code $LASTEXITCODE."
        }
    }
    Add-BenchmarkEngine PSWriteOffice {
        foreach ($operation in 'Markdown', 'Html') {
            Add-BenchmarkOperation $operation {
                param($case, $run)
                $null = PSWriteOffice\Convert-OfficeDocument -LiteralPath $case.InputPath -OutputPath $run.OutputPath -FailOnLoss
            }
        }
        Add-BenchmarkOperation ReadCells {
            param($case, $run)
            $notebook = PSWriteOffice\Open-OfficeNotebook -LiteralPath $case.InputPath
            try {
                $run.Cells = @($notebook | PSWriteOffice\Get-OfficeNotebookCell | ForEach-Object {
                    [pscustomobject]@{ Kind = [string]$_.Kind; Source = $_.Source }
                })
            } finally {
                $notebook.Dispose()
            }
        }
        Add-BenchmarkOperation StartupMarkdown $startup
        Add-BenchmarkOperation BatchMarkdown {
            param($case, $run)
            $null = PSWriteOffice\Convert-OfficeDocument -LiteralPath $run.BatchInputs -To .md -OutputDirectory $run.BatchOutput -ThrottleLimit $case.ThrottleLimit -FailOnLoss
        }
    }
    if ($peerManifest) {
        Add-BenchmarkEngine PowerShellNotebook {
            foreach ($operation in 'Markdown', 'Html') {
                Add-BenchmarkOperation $operation {
                    param($case, $run)
                    $format = @{}
                    if ($case.Operation -eq 'Html') {
                        $format.AsHTML = $true
                    }
                    $null = PowerShellNotebook\ConvertFrom-IPYNB -Path $case.InputPath -Destination $run.OutputPath -IncludeOutput @format
                }
            }
            Add-BenchmarkOperation ReadCells {
                param($case, $run)
                $run.Cells = @(PowerShellNotebook\Get-NotebookContent -Path $case.InputPath | ForEach-Object {
                    [pscustomobject]@{ Kind = $_.Type; Source = $_.Source }
                })
            }
            Add-BenchmarkOperation StartupMarkdown $startup
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)

        if ($case.Operation -eq 'ReadCells') {
            Assert-NotebookPerformanceCells $case $run.Cells
            return
        }
        if ($case.Operation -eq 'BatchMarkdown') {
            $files = @(Get-ChildItem -LiteralPath $run.BatchOutput -Filter '*.md' -File -Recurse)
            Assert-BenchmarkValue -Actual $files.Count -Expected $case.Files
            for ($index = 1; $index -le $case.Files; $index++) {
                $path = Join-Path $run.BatchOutput "notebook-$index.md"
                Assert-NotebookPerformanceArtifact $path $case.Cells -BatchIndex $index
            }
            return
        }
        Assert-NotebookPerformanceArtifact $run.OutputPath $case.Cells -Html:($case.Operation -eq 'Html')
        if ($case.Operation -eq 'StartupMarkdown' -and $observeMemory) {
            $observation = if ($darwinTime) {
                [regex]::Match([IO.File]::ReadAllText($run.TimeLog), '(?m)^\s*(\d+)\s+maximum resident set size')
            } else {
                [regex]::Match(($run.ChildOutput -join "`n"), '(?m)^WorkerPeakWorkingSetBytes=(\d+)')
            }
            if ($observation.Success) {
                $run.WorkerPeakWorkingSetBytes = [long]$observation.Groups[1].Value
            }
        }
    }
    Add-BenchmarkMetric VerifiedCells { param($case, $run) $case.Cells * $case.Files }
    Add-BenchmarkMetric VerifiedFiles { param($case, $run) $case.Files }
    if ($mode -eq 'Cold' -and $observeMemory) {
        Add-BenchmarkMetric WorkerPeakWorkingSetBytes { param($case, $run) $run.WorkerPeakWorkingSetBytes }
    }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
