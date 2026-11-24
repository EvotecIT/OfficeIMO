[CmdletBinding()]
param(
    [ValidateSet('Codec', 'Projection', 'Export', 'Fidelity')] [string] $Family = 'Codec',
    [ValidateSet('OfficeIMO', 'Reference', 'Microsoft', 'Verso')] [string[]] $Engine,
    [string[]] $Case,
    [ValidateSet('Parse', 'RoundTrip', 'SourceProjection', 'RichRoundTrip', 'Html', 'Markdown', 'Pdf')] [string[]] $Operation,
    [ValidateRange(2, 10000)] [int[]] $Cells = @(12, 120, 1200),
    [ValidateRange(0, 8192)] [int] $CodeLineCharacters = 32,
    [ValidateRange(0, 8192)] [int] $OutputCharacters = 32,
    [ValidateRange(16, 2048)] [int] $RichImagePixels = 64,
    [ValidateRange(0, 100)] [int] $WarmupCount = 2,
    [ValidateRange(1, 100)] [int] $IterationCount = 5,
    [ValidateRange(10, 240)] [int] $WorkerTimeoutSeconds = 60,
    [string] $ManagedBinary = (Join-Path $PSScriptRoot '../../OfficeIMO.Notebook.Benchmarks/bin/Release/net10.0/OfficeIMO.Notebook.Benchmarks.dll'),
    [string] $OutputRoot = (Join-Path $PSScriptRoot '../../Ignore/Benchmarks/Notebook'),
    [string] $ModulePath = 'PSPublishModule',
    [string] $PythonPath = 'python3',
    [string] $DotNetPath = 'dotnet',
    [string] $PdfTextTool = 'pdftotext',
    [string] $PdfImageTool = 'pdfimages',
    [Nullable[UInt64]] $ProcessorAffinityMask,
    [ValidateSet('Idle', 'BelowNormal', 'Normal', 'AboveNormal', 'High')] [string] $ProcessPriority = 'Normal',
    [switch] $Cold,
    [switch] $Plan
)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
if (@($Cells | Sort-Object -Unique).Count -ne $Cells.Count) { throw 'Cell counts must be distinct.' }
if ($Cold -and $Family -eq 'Fidelity') { throw 'Select Codec, Projection, or Export to measure cold worker startup.' }

$families = @{
    Codec = @{ Engines = @('OfficeIMO', 'Reference'); Operations = @('Parse', 'RoundTrip') }
    Projection = @{ Engines = @('OfficeIMO', 'Microsoft', 'Verso', 'Reference'); Operations = @('SourceProjection') }
    Export = @{ Engines = @('OfficeIMO', 'Reference'); Operations = @('Html', 'Markdown', 'Pdf') }
    Fidelity = @{ Engines = @('OfficeIMO', 'Microsoft', 'Verso', 'Reference'); Operations = @('RichRoundTrip') }
}
$selection = $families[$Family]
if (-not $Engine) { $Engine = $selection.Engines }
if (-not $Operation) { $Operation = $selection.Operations }
if (@($Engine | Where-Object { $_ -notin $selection.Engines }).Count -gt 0 -or
    @($Operation | Where-Object { $_ -notin $selection.Operations }).Count -gt 0) {
    throw 'Choose engines and operations that belong to the selected benchmark family.'
}

$managed = @{}
$reference = @{}
$dotnet = ''
$python = ''
$pdfText = ''
$pdfImages = ''
$pythonVersion = ''
if (-not $Plan) {
    $python = (Get-Command $PythonPath -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    $referenceArguments = @((Join-Path $PSScriptRoot 'notebook-reference.py'), 'environment')
    if ('Pdf' -in $Operation -and 'Reference' -in $Engine) { $referenceArguments += '--probe-chromium' }
    $pythonVersion = & $python @referenceArguments
    $reference = $pythonVersion | ConvertFrom-Json -AsHashtable
    if ('Pdf' -in $Operation) {
        $pdfText = (Get-Command $PdfTextTool -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1).Source
        $pdfImages = (Get-Command $PdfImageTool -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1).Source
    }
    if (@($Engine | Where-Object { $_ -ne 'Reference' }).Count -gt 0 -and (Test-Path -LiteralPath $ManagedBinary -PathType Leaf)) {
        $ManagedBinary = (Resolve-Path -LiteralPath $ManagedBinary).Path
        $dotnet = (Get-Command $DotNetPath -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
        $managed = & $dotnet $ManagedBinary environment | ConvertFrom-Json -AsHashtable
    }
}

$session = [Guid]::NewGuid().ToString('N')
$OutputRoot = [IO.Path]::GetFullPath((Join-Path $OutputRoot $Family))
Import-Module (Join-Path $PSScriptRoot 'notebook-worker.psm1') -ErrorAction Stop
Import-Module $ModulePath -ErrorAction Stop
$parameters = @{
    Path = Join-Path $PSScriptRoot 'notebook.benchmark.ps1'
    OutputRoot = $OutputRoot
    Variable = @{
        Family = $Family; Cells = ($Cells -join ','); CodeLineCharacters = $CodeLineCharacters
        OutputCharacters = $OutputCharacters; RichImagePixels = $RichImagePixels; Session = $session
        PythonPath = $python; DotNetPath = $dotnet; ManagedBinary = $ManagedBinary
        PdfText = $pdfText; PdfImages = $pdfImages; PythonEnvironment = $pythonVersion
        HasOfficeIMO = [bool]$managed.ok; HasMicrosoft = $managed.microsoft -eq '1.0.0-beta.26120.1'
        HasVerso = $managed.verso -eq '1.2.3'; ManagedRuntime = $managed.runtime
        HasNbformat = $reference.nbformat -eq '5.11.1' -and $reference.fastjsonschema -eq '2.22.3' -and $reference.validator -eq 'fastjsonschema'
        HasNbconvert = $reference.nbconvert -eq '7.17.2'
        HasChromium = [bool]$reference.chromium -and $reference.playwright -eq '1.63.0'; Planning = [bool]$Plan
        Cold = [bool]$Cold
        InputsRoot = Join-Path $OutputRoot 'inputs'
        WorkerTimeoutSeconds = $WorkerTimeoutSeconds
    }
    Engine = $Engine; Operation = $Operation; WarmupCount = $WarmupCount; IterationCount = $IterationCount
    Plan = $Plan
}
if ($Case) { $parameters.Case = $Case }
if ($null -ne $ProcessorAffinityMask) { $parameters.ProcessorAffinityMask = [UInt64]$ProcessorAffinityMask }
if ($IsWindows) { $parameters.ProcessPriority = $ProcessPriority }
if ($Family -eq 'Fidelity') {
    $parameters.WarmupCount = 0
    $parameters.IterationCount = 1
    $parameters.RunMode = 'fidelity-only'
}
if ($Cold) {
    $parameters.WarmupCount = 0
    $parameters.RunMode = 'cold-worker'
}

try {
    $result = Invoke-BenchmarkSuite @parameters
    $result
    if (-not $Plan) {
        $failed = @($result.Samples | Where-Object Status -eq 'Failed')
        if ($failed.Count -gt 0) {
            $failure = $failed[0]
            throw "Notebook benchmark failed ($($failure.Engine)/$($failure.Operation)/$($failure.Scenario)). $($failure.Reason)"
        }
        if (@($result.Samples | Where-Object Status -eq 'Succeeded').Count -eq 0) {
            throw 'No benchmark lanes could run. Inspect skipped lanes and the prerequisites.'
        }
    }
} finally {
    Stop-NotebookWorkers -Session $session
}
