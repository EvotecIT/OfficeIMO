$mode = Get-BenchmarkInput Mode -Required
$scenarios = (Get-BenchmarkInput Scenarios -Required) -split ','
$fixtureRoot = Get-BenchmarkInput FixtureRoot -Required
$moduleManifest = Get-BenchmarkInput ModuleManifest ''
$peerManifest = Get-BenchmarkInput PowerShellNotebookManifest ''
$powerShellPath = Get-BenchmarkInput PowerShellPath -Required
$dotnetRepl = Get-BenchmarkInput DotnetReplPath ''
$dotnetReplVersion = Get-BenchmarkInput DotnetReplVersion ''
$officeimoWarmModuleLoaded = Get-BenchmarkInput OfficeimoWarmModuleLoaded $false -Bool
$interactiveAssembly = Get-BenchmarkInput InteractiveBenchmarkAssembly ''
$childScript = Join-Path $PSScriptRoot 'Invoke-NotebookExecutionChild.ps1'
. (Join-Path $PSScriptRoot 'notebook-execution.fixtures.ps1')

if ($officeimoWarmModuleLoaded) {
    # Keep the product assembly in the module context that owns its dependencies.
    $command = Get-Command PSWriteOffice\New-OfficeNotebook -ListImported
    $context = [System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext($command.ImplementingType.Assembly)
    $officeimoAssembly = $context.LoadFromAssemblyName([Reflection.AssemblyName]::new('OfficeIMO.Notebook.PowerShell'))
    $officeimoSessionType = $officeimoAssembly.GetType('OfficeIMO.Notebook.PowerShell.PowerShellNotebookSession', $true)
    $officeimoSettingsType = $officeimoAssembly.GetType('OfficeIMO.Notebook.PowerShell.PowerShellNotebookSessionOptions', $true)
    $officeimoIsolationType = $officeimoAssembly.GetType('OfficeIMO.Notebook.PowerShell.PowerShellNotebookIsolation', $true)
    $officeimoNotebookAssembly = $context.LoadFromAssemblyName([Reflection.AssemblyName]::new('OfficeIMO.Notebook'))
    $officeimoNotebookType = $officeimoNotebookAssembly.GetType('OfficeIMO.Notebook.NotebookDocument', $true)
}
$fixtures = @{}
Get-NotebookExecutionCases $scenarios $fixtureRoot | ForEach-Object { $fixtures[$_.Name] = $_ }

function Close-NotebookBenchmarkSession {
    param($Case, $Run)

    if ($Run.PSObject.Properties['Session'] -and $null -ne $Run.Session) {
        if ($Case.Engine -eq 'PowerShellNotebook') {
            $Run.Session.Close()
            $Run.Session.PowerShell.Dispose()
            $Run.Session.Runspace.Dispose()
        } else {
            $Run.Session.Dispose()
        }
        $Run.Session = $null
    }
    if ($Run.PSObject.Properties['Notebook'] -and $null -ne $Run.Notebook) {
        $Run.Notebook.Dispose()
        $Run.Notebook = $null
    }
}

New-BenchmarkSuite 'notebook-execution' {
    Set-BenchmarkPolicy -Warmup 0 -Iteration 5 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Execute ordered PowerShell cells in one session. Verify each saved or captured text result against independent state markers and verify the UTF-8 checksum.'
    Add-BenchmarkMetadata Timing 'Cold includes a fresh process, notebook read, session start, execution, and saved artifact. Warm starts and primes the session outside timing, then executes and captures the batch. Session disposal and validation are outside timing.'
    Add-BenchmarkMetadata PowerShell $PSVersionTable.PSVersion.ToString()
    Add-BenchmarkMetadata Framework ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    Add-BenchmarkMetadata PowerShellNotebook 'PowerShell Gallery 3.0.0, MIT'
    Add-BenchmarkMetadata DotNetInteractive 'PowerShell 1.0.0-beta.26120.1, MIT'
    if ($dotnetReplVersion) {
        Add-BenchmarkMetadata DotNetRepl "$dotnetReplVersion, MIT, embedded PowerShell SDK 7.5.4"
    }
    if ($moduleManifest) {
        Add-BenchmarkMetadata PSWriteOfficeManifestSha256 (Get-FileHash -LiteralPath $moduleManifest).Hash
    }
    if ($officeimoWarmModuleLoaded) {
        Add-BenchmarkMetadata OfficeimoAssemblySha256 (Get-FileHash -LiteralPath $officeimoAssembly.Location).Hash
        $target = $officeimoAssembly.GetCustomAttributes([Runtime.Versioning.TargetFrameworkAttribute], $false)
        Add-BenchmarkMetadata OfficeimoTargetFramework $target[0].FrameworkName
    }
    if ($interactiveAssembly) {
        Add-BenchmarkMetadata InteractiveBenchmarkAssemblySha256 (Get-FileHash -LiteralPath $interactiveAssembly).Hash
    }
    if ($peerManifest) {
        Add-BenchmarkMetadata PowerShellNotebookManifestSha256 (Get-FileHash -LiteralPath $peerManifest).Hash
        $peerSource = Join-Path (Split-Path $peerManifest) 'InvokePowerShellNotebook.ps1'
        Add-BenchmarkMetadata PowerShellNotebookExecutionSha256 (Get-FileHash -LiteralPath $peerSource).Hash
    }

    Add-BenchmarkCaseSource {
        foreach ($name in $scenarios) {
            $fixture = $fixtures[$name]
            $inputHash = if (Test-Path -LiteralPath $fixture.InputPath) {
                (Get-FileHash -LiteralPath $fixture.InputPath).Hash
            } else {
                'planned'
            }
            [pscustomobject]@{
                Name = $name
                CodeCells = $fixture.Count
                TextBytes = $fixture.ExpectedBytes
                TextSha256 = $fixture.ExpectedSha256
                NotebookSha256 = $inputHash
            }
        }
    }

    Set-BenchmarkSetup {
        param($case, $run)
        $run.Fixture = $fixtures[$case.Scenario]
        $run.OutputPath = Join-Path $run.OutputDirectory 'executed.ipynb'
        if ($mode -eq 'Cold') {
            if ([IO.File]::Exists($run.OutputPath)) {
                [IO.File]::Delete($run.OutputPath)
            }
            $manifest = if ($case.Engine -like 'OfficeIMO*') { $moduleManifest } else { $peerManifest }
            $run.PowerShellArguments = @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', $childScript,
                '-Engine', $case.Engine, '-ModuleManifest', $manifest, '-InputPath', $run.Fixture.InputPath,
                '-OutputPath', $run.OutputPath, '-PowerShellPath', $powerShellPath)
            return
        }

        $primeCode = "'notebook-benchmark-ready'"
        try {
            switch ($case.Engine) {
                { $_ -in 'OfficeIMO', 'OfficeIMOProcess' } {
                    $settings = [Activator]::CreateInstance($officeimoSettingsType)
                    if ($case.Engine -eq 'OfficeIMO') {
                        $settings.Isolation = [Enum]::Parse($officeimoIsolationType, 'Runspace')
                    } else {
                        $settings.ExecutablePath = $powerShellPath
                    }
                    $run.Session = $officeimoSessionType::StartAsync($settings).GetAwaiter().GetResult()
                    $prime = $officeimoNotebookType::Create()
                    try {
                        $null = $prime.AddCodeCell($primeCode, 'PowerShell')
                        $primed = $run.Session.ExecuteAsync($prime).GetAwaiter().GetResult()
                        if (-not $primed.Succeeded) { throw 'OfficeIMO PowerShell session did not prime.' }
                    } finally {
                        $prime.Dispose()
                    }
                    $run.Notebook = $officeimoNotebookType::Create()
                    foreach ($source in $run.Fixture.Sources) {
                        $null = $run.Notebook.AddCodeCell($source, 'PowerShell')
                    }
                }
                'PowerShellNotebook' {
                    $run.Session = PowerShellNotebook\New-PSNotebookRunspace
                    $null = $run.Session.Invoke($primeCode)
                    if ($run.Session.PowerShell.Streams.Error.Count -ne 0) {
                        throw 'PowerShellNotebook runspace did not prime.'
                    }
                }
                'DotNetInteractive' {
                    $run.Session = [NotebookExecution.Interactive.Benchmarks.InteractiveBenchmarkSession]::new()
                    $null = $run.Session.ExecuteAsync([string[]] @($primeCode)).GetAwaiter().GetResult()
                }
            }
        } catch {
            Close-NotebookBenchmarkSession $case $run
            throw
        }
    }


    $runPowerShellFile = {
        param($case, $run)
        $arguments = $run.PowerShellArguments
        $null = & $powerShellPath @arguments
        if ($LASTEXITCODE -ne 0) {
            throw "$($case.Engine) exited with code $LASTEXITCODE."
        }
    }

    Add-BenchmarkEngine OfficeIMO {
        Add-BenchmarkOperation RunAndSave $runPowerShellFile
        Add-BenchmarkOperation ExecuteCells {
            param($case, $run)
            try {
                $run.Report = $run.Session.ExecuteAsync($run.Notebook).GetAwaiter().GetResult()
            } catch {
                Close-NotebookBenchmarkSession $case $run
                throw
            }
        }
    }
    Add-BenchmarkEngine OfficeIMOProcess {
        Add-BenchmarkOperation RunAndSave $runPowerShellFile
        Add-BenchmarkOperation ExecuteCells {
            param($case, $run)
            try {
                $run.Report = $run.Session.ExecuteAsync($run.Notebook).GetAwaiter().GetResult()
            } catch {
                Close-NotebookBenchmarkSession $case $run
                throw
            }
        }
    }
    Add-BenchmarkEngine PowerShellNotebook {
        Add-BenchmarkOperation RunAndSave $runPowerShellFile
        Add-BenchmarkOperation ExecuteCells {
            param($case, $run)
            try {
                $run.Text = @(foreach ($source in $run.Fixture.Sources) {
                    $run.Session.PowerShell.Commands.Clear()
                    $run.Session.PowerShell.Streams.Error.Clear()
                    $text = $run.Session.Invoke($source)
                    if ($run.Session.PowerShell.Streams.Error.Count -ne 0) { throw 'PowerShellNotebook emitted an error.' }
                    -join $text
                })
            } catch {
                Close-NotebookBenchmarkSession $case $run
                throw
            }
        }
    }
    Add-BenchmarkEngine DotNetInteractive {
        Add-BenchmarkOperation RunAndSave {
            param($case, $run)
            $null = & $dotnetRepl --default-kernel pwsh --run $run.Fixture.InputPath --output-path $run.OutputPath --output-format ipynb --exit-after-run
            if ($LASTEXITCODE -ne 0) { throw "dotnet-repl exited with code $LASTEXITCODE." }
        }
        Add-BenchmarkOperation ExecuteCells {
            param($case, $run)
            try {
                $run.Text = $run.Session.ExecuteAsync($run.Fixture.Sources).GetAwaiter().GetResult()
            } catch {
                Close-NotebookBenchmarkSession $case $run
                throw
            }
        }
    }

    Add-BenchmarkValidation {
        param($case, $run)
        try {
            $text = if ($mode -eq 'Cold') {
                @(Read-NotebookExecutionOutput $run.OutputPath $run.Fixture.Count)
            } elseif ($case.Engine -like 'OfficeIMO*') {
                if (-not $run.Report.Succeeded -or ($run.Report.Cells | Where-Object OutputTruncated)) {
                    throw 'OfficeIMO reported an unsuccessful execution or truncated output.'
                }
                @(foreach ($cell in $run.Report.Notebook.Cells) {
                    $parts = foreach ($output in $cell.Outputs) {
                        if ($output.Kind -eq 'Stream' -and $output.Name -eq 'stdout') {
                            $output.Text
                        } elseif ($output.Kind -in 'ExecuteResult', 'DisplayData' -and $output.Data.ContainsKey('text/plain')) {
                            $output.Data['text/plain'].GetString()
                        } else {
                            throw "Unexpected OfficeIMO cell output '$($output.Kind)/$($output.Name)'."
                        }
                    }
                    -join $parts
                })
            } else {
                $run.Text
            }
            Assert-NotebookExecutionText $run.Fixture $text
        } finally {
            Close-NotebookBenchmarkSession $case $run
        }
    }
    Add-BenchmarkMetric ExecutedCells { param($case, $run) $run.Fixture.Count }
    Add-BenchmarkMetric ValidatedTextBytes { param($case, $run) $run.Fixture.ExpectedBytes }
    Add-BenchmarkComparison Engine -Baseline OfficeIMO -Metric MedianMs -TieTolerance 0.05
}
