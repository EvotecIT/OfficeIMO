# PowerShell notebook execution benchmark

The suite compares the OfficeIMO, PowerShellNotebook, and .NET Interactive execution paths on ordered, deterministic PowerShell cells. The first cell defines a function and a counter. Subsequent cells call the function and advance the counter in the same session. The large text case adds 512 lines to the final output.

PowerForge rotates engine order and validates each sample. Cold validation reopens the saved notebook and verifies its introduction, cell order, output types, exact text, and a checksum. Warm validation checks captured text against the same expected results. The validator allows terminal line endings added by the engines' text formatters.

## Cold notebook execution

From the repository root, build PSWriteOffice or locate the manifest of a built module. Run these commands with PowerShell 7.5.4 and a .NET 10 SDK. Use a task-owned output directory on the development volume for the third-party tools, build output, and measurements.

~~~powershell
$scratch = Join-Path $PWD ('Ignore/Benchmarks/NotebookExecution/' + [guid]::NewGuid().ToString('N'))
$peerRoot = Join-Path $scratch 'modules'
New-Item -ItemType Directory -Path $peerRoot -Force | Out-Null
Save-Module -Name PowerShellNotebook -RequiredVersion 3.0.0 -Path $peerRoot
$replRoot = Join-Path $scratch 'tools'
dotnet tool install dotnet-repl --version 0.3.259 --tool-path $replRoot
$repl = Join-Path $replRoot $(if ($IsWindows) { 'dotnet-repl.exe' } else { 'dotnet-repl' })
$comparison = Join-Path $PWD 'PowerShell/PSWriteOffice/Benchmarks/Compare-NotebookExecution.ps1'
$cold = @{
    Mode = 'Cold'
    PowerShellNotebookManifest = Join-Path $peerRoot 'PowerShellNotebook/3.0.0/PowerShellNotebook.psd1'
    DotnetReplPath = $repl
    OutputRoot = Join-Path $scratch 'cold'
}
& $comparison @cold -Plan
& $comparison @cold
~~~

Cold timing includes process startup, loading the notebook, opening one session, executing its cells, saving the `.ipynb`, and process exit. The suite calls `Invoke-OfficeNotebook`, `Invoke-PowerShellNotebook`, and `dotnet-repl`. The REPL hosts .NET Interactive PowerShell `1.0.0-beta.26120.1` with PowerShell SDK 7.5.4. [PowerShellNotebook 3.0.0](https://www.powershellgallery.com/packages/PowerShellNotebook/3.0.0) · [dotnet-repl 0.3.259](https://www.nuget.org/packages/dotnet-repl/0.3.259)

## Warm session execution

Run the warm lane in PowerShell 7.6 on .NET 10. Set the session's `$scratch` variable to the same directory used above. Build the .NET Interactive benchmark adapter from the repository root. Use the same PSWriteOffice manifest as the cold lane.

~~~powershell
$adapter = 'PowerShell/PSWriteOffice/Benchmarks/NotebookExecution/NotebookExecution.Interactive.Benchmarks/NotebookExecution.Interactive.Benchmarks.csproj'
$adapterArtifacts = Join-Path $scratch 'interactive-adapter'
dotnet build $adapter -c Release --artifacts-path $adapterArtifacts
$adapterOut = Join-Path $adapterArtifacts 'bin/NotebookExecution.Interactive.Benchmarks/release'
$comparison = Join-Path $PWD 'PowerShell/PSWriteOffice/Benchmarks/Compare-NotebookExecution.ps1'
$warm = @{
    Mode = 'Warm'
    PowerShellNotebookManifest = Join-Path $scratch 'modules/PowerShellNotebook/3.0.0/PowerShellNotebook.psd1'
    InteractiveBenchmarkAssembly = Join-Path $adapterOut 'NotebookExecution.Interactive.Benchmarks.dll'
    OutputRoot = Join-Path $scratch 'warm'
}
& $comparison @warm -Plan
& $comparison @warm
~~~

Warm setup opens and primes a fresh engine session per sample. Timing covers submission of every cell and capture of its output. Validation and session disposal follow timing. OfficeIMO uses `PowerShellNotebookSession.ExecuteAsync`, PowerShellNotebook uses `New-PSNotebookRunspace`, and the adapter submits cells to .NET Interactive's `PowerShellKernel`. OfficeIMO calls the public .NET API from the PSWriteOffice module payload and uses the same PowerShell executable as the benchmark host. The adapter references the pinned MIT kernel with serviced cryptography packages. [Exact .NET Interactive PowerShell package](https://www.nuget.org/packages/Microsoft.DotNet.Interactive.PowerShell/1.0.0-beta.26120.1)

Use `-Scenario`, `-Engine`, and `-IterationCount` to narrow a diagnostic run. Start with a single iteration, then run the balanced default matrix. Review each run's `summary.json`, `metadata.json`, and raw `samples.json`. Read cold and warm medians independently. PowerForge allocation metrics describe the supervising process. Elapsed time captures work done in that process and any engine workers.
