#Requires -Version 7.4
<#
.SYNOPSIS
Compares PSWriteOffice's Markdown-to-HTML path with PowerShell's ConvertFrom-Markdown.

.DESCRIPTION
Times `Convert-OfficeDocument -Text <markdown> -From Markdown -To Html` against
`ConvertFrom-Markdown -InputObject <markdown>` on the same inputs, and splits the
PSWriteOffice time into the cmdlet layer, the converter, parser setup, parsing,
rendering and the HTML output wrapper.

Warm numbers are medians of several rounds, each a loop of calls in one process, after
a time-based warm-up (WarmupMilliseconds per lane) so the engine code has reached its
optimized JIT tier.
The first call is measured in a fresh process per tool and reported separately.
Every lane runs through the PowerShell pipeline, so loop overhead is included in
each lane equally; compare lanes against each other rather than reading them as
raw engine timings. The BenchmarkDotNet project OfficeIMO.Markdown.Benchmarks
remains the engine comparison against standalone Markdig.

The module is loaded from this checkout's Release development binaries
(Sources/PSWriteOffice/bin/Release/net8.0), so build it first. Every lane's output is
checked for the input's text before any timing starts.

    dotnet build Sources/PSWriteOffice/PSWriteOffice.csproj -c Release -f net8.0

.EXAMPLE
pwsh -NoLogo -NoProfile -File ./Benchmarks/Compare-MarkdownPerformance.ps1

.EXAMPLE
pwsh -NoLogo -NoProfile -File ./Benchmarks/Compare-MarkdownPerformance.ps1 -Rounds 9 -Iterations 500 -SkipFirstCall
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 100)][int] $Rounds = 7,
    [ValidateRange(1, 100000)][int] $Iterations = 300,
    [ValidateRange(0, 60000)][int] $WarmupMilliseconds = 1500,
    [ValidateSet('Paragraph', 'SmallReadme', 'MediumDocument')][string[]] $Corpus = @('Paragraph', 'SmallReadme', 'MediumDocument'),
    [switch] $SkipFirstCall,
    [switch] $AsObject
)

$ErrorActionPreference = 'Stop'
$moduleManifest = Join-Path (Split-Path $PSScriptRoot -Parent) 'PSWriteOffice.psd1'

$readmeSection = @'
# OfficeIMO Markdown Overview

OfficeIMO.Markdown can build, parse, inspect, and render Markdown for chat, docs, and reports.

## Features

- Fluent builders for headings, paragraphs, tables, lists, and fenced code
- Typed document queries for headings, list items, top-level blocks, and descendants
- HTML rendering with anchor links and table-of-contents support

## Sample Table

| Area | Status | Notes |
| --- | --- | --- |
| Reader | Active | Typed document traversal |
| Renderer | Active | Portable HTML output |
| Docs | Improving | More examples and contract docs |

## Example

```csharp
var doc = MarkdownDocument.Create();
doc.H1("Status");
doc.P("Everything is working.");
```

Additional notes:

1. Keep README output readable.
2. Keep parser behavior predictable.
3. Keep portability in mind.

> This quote keeps the corpus close to a normal README.

'@ -replace "`r`n", "`n"

$inputs = [ordered]@{
    Paragraph      = 'Deployment finished in **staging**; see the [runbook](https://example.com/runbook) for the next step.'
    SmallReadme    = $readmeSection
    MediumDocument = (@($readmeSection) * 12) -join "`n"
}

function Measure-Median {
    param([scriptblock] $Operation, [int] $Rounds, [int] $Iterations)
    # Warm for a fixed time rather than a call count, so the measured code has reached its optimized tier.
    $warm = [Diagnostics.Stopwatch]::StartNew()
    while ($warm.ElapsedMilliseconds -lt $WarmupMilliseconds) { $null = & $Operation }
    $samples = for ($round = 0; $round -lt $Rounds; $round++) {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        for ($i = 0; $i -lt $Iterations; $i++) { $null = & $Operation }
        $watch.Elapsed.TotalMilliseconds * 1000 / $Iterations
    }
    $sorted = @($samples | Sort-Object)
    $sorted[[int][Math]::Floor($sorted.Count / 2)]
}

# The development loader prefers Debug binaries unless told otherwise; child processes inherit this.
$env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = 'Release'
Import-Module $moduleManifest -Force
# The engine assemblies load on first use.
$null = Convert-OfficeDocument -Text 'warm' -From Markdown -To Html
$engine = @{}
foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
    $name = $assembly.GetName().Name
    if ($name -in 'OfficeIMO', 'OfficeIMO.Core', 'OfficeIMO.Markdown') { $engine[$name] = $assembly }
}
if ($engine.Count -ne 3) { throw 'The PSWriteOffice engine assemblies did not load.' }
if ($engine['OfficeIMO.Markdown'].Location -notmatch '[\\/]Release[\\/]') {
    throw "The module loaded $($engine['OfficeIMO.Markdown'].Location); build the Release binaries first."
}
$null = ConvertFrom-Markdown -InputObject 'warm'
$markdigAssembly = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Markdig.Signed' } | Select-Object -First 1
$markdigVersion = $markdigAssembly.GetName().Version

# Engine stages are timed in compiled loops: calling .NET methods from a PowerShell loop would mostly
# measure PowerShell's method binder. The helper binds to the engine assemblies the module loaded.
$stageSource = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using OfficeIMO;
using OfficeIMO.Markdown;
using Markdig;

public static class OfficeIMOMarkdownStages {
    private static readonly UTF8Encoding Encoding = new UTF8Encoding(false, true);
    private static readonly Markdig.MarkdownPipeline AdvancedPipeline = new Markdig.MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

    // The converter's Markdown route parses without a syntax tree and renders an HTML fragment.
    private static MarkdownLoadOptions Load() => new MarkdownLoadOptions { CaptureSyntaxTree = false };
    private static MarkdownHtmlExportOptions Fragment() => new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Fragment };

    public static string[] Names => new[] { "Converter call (stream in, stream out)", "Parser setup (empty input)", "Parse", "Render HTML fragment", "Render to stream (HTML wrapper)", "Markdig ToHtml, advanced extensions" };

    // The HTML each engine lane produces, so the script can check it before timing anything.
    public static string[] Outputs(string markdown) {
        MarkdownDocument parsed = MarkdownDocument.Parse(markdown, Load());
        using var source = new MemoryStream(Encoding.GetBytes(markdown), false);
        using var converted = new MemoryStream();
        var options = new OfficeConversionOptions { SourceName = "source.md", TargetName = "output.html" };
        OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, converted, OfficeDocumentFormat.Html, options).GetAwaiter().GetResult();
        using var streamed = new MemoryStream();
        parsed.SaveAsHtmlAsync(streamed, Fragment()).GetAwaiter().GetResult();
        return new[] {
            Encoding.GetString(converted.ToArray()),
            parsed.ToHtml(Fragment()),
            Encoding.GetString(streamed.ToArray()),
            Markdig.Markdown.ToHtml(markdown, AdvancedPipeline)
        };
    }

    public static double[] Measure(string markdown, int rounds, int iterations, int warmupMilliseconds) {
        MarkdownDocument parsed = MarkdownDocument.Parse(markdown, Load());
        var stages = new Action[] {
            () => {
                using var source = new MemoryStream(Encoding.GetBytes(markdown), false);
                using var target = new MemoryStream();
                var options = new OfficeConversionOptions { SourceName = "source.md", TargetName = "output.html" };
                OfficeConverter.ConvertAsync(source, OfficeDocumentFormat.Markdown, target, OfficeDocumentFormat.Html, options).GetAwaiter().GetResult();
                Encoding.GetString(target.ToArray());
            },
            () => MarkdownDocument.Parse(string.Empty, Load()),
            () => MarkdownDocument.Parse(markdown, Load()),
            () => parsed.ToHtml(Fragment()),
            () => {
                using var target = new MemoryStream();
                parsed.SaveAsHtmlAsync(target, Fragment()).GetAwaiter().GetResult();
            },
            () => Markdig.Markdown.ToHtml(markdown, AdvancedPipeline)
        };
        var medians = new double[stages.Length];
        for (int stage = 0; stage < stages.Length; stage++) {
            var warm = Stopwatch.StartNew();
            while (warm.ElapsedMilliseconds < warmupMilliseconds) stages[stage]();
            var samples = new List<double>();
            for (int round = 0; round < rounds; round++) {
                var watch = Stopwatch.StartNew();
                for (int i = 0; i < iterations; i++) stages[stage]();
                samples.Add(watch.Elapsed.TotalMilliseconds * 1000 / iterations);
            }
            samples.Sort();
            medians[stage] = samples[samples.Count / 2];
        }
        return medians;
    }
}
'@
$references = @($engine.Values | ForEach-Object Location) + $markdigAssembly.Location + @(
    'netstandard', 'System.Runtime', 'System.Collections', 'System.Text.Encoding.Extensions', 'System.Threading.Tasks', 'System.Runtime.Extensions')
# The module loads its engine in its own load context; the helper uses those same loaded copies.
$moduleContext = [Runtime.Loader.AssemblyLoadContext]::GetLoadContext($engine['OfficeIMO.Markdown'])
$moduleAssemblies = @{}
foreach ($assembly in $moduleContext.Assemblies) { $moduleAssemblies[$assembly.GetName().Name] = $assembly }
[Runtime.Loader.AssemblyLoadContext]::Default.add_Resolving([Func[Runtime.Loader.AssemblyLoadContext, Reflection.AssemblyName, Reflection.Assembly]]{
    param($context, $assemblyName)
    $moduleAssemblies[$assemblyName.Name]
})
if (-not ('OfficeIMOMarkdownStages' -as [type])) {
    Add-Type -TypeDefinition $stageSource -ReferencedAssemblies $references -IgnoreWarnings -WarningAction SilentlyContinue
}

# Text every lane's HTML must contain for its input, so an empty or partial conversion cannot pass as a faster one.
$expectedText = @{
    Paragraph      = @('Deployment finished in', 'staging', 'runbook', 'for the next step.')
    SmallReadme    = @('OfficeIMO Markdown Overview', 'Typed document traversal', 'MarkdownDocument.Create();', 'Keep portability in mind.', 'close to a normal README.')
    MediumDocument = @('OfficeIMO Markdown Overview', 'Typed document traversal', 'MarkdownDocument.Create();', 'Keep portability in mind.', 'close to a normal README.')
}

function Assert-LaneOutput {
    param([string] $Lane, [string] $InputName, [string] $Html)
    $text = [Net.WebUtility]::HtmlDecode($Html)
    foreach ($expected in $expectedText[$InputName]) {
        if ($text.IndexOf($expected, [StringComparison]::Ordinal) -lt 0) {
            throw "$Lane did not produce the expected HTML for $InputName (missing '$expected')."
        }
    }
}

foreach ($inputName in $Corpus) {
    $markdown = $inputs[$inputName]
    Assert-LaneOutput 'Convert-OfficeDocument' $inputName ([string](Convert-OfficeDocument -Text $markdown -From Markdown -To Html))
    Assert-LaneOutput 'ConvertFrom-Markdown' $inputName (ConvertFrom-Markdown -InputObject $markdown).Html
    $engineOutputs = [OfficeIMOMarkdownStages]::Outputs($markdown)
    $engineLanes = 'converter call', 'render HTML fragment', 'render to stream', 'Markdig ToHtml'
    for ($index = 0; $index -lt $engineLanes.Count; $index++) { Assert-LaneOutput $engineLanes[$index] $inputName $engineOutputs[$index] }
}

$results = foreach ($inputName in $Corpus) {
    $markdown = $inputs[$inputName]
    $lanes = [ordered]@{
        'Convert-OfficeDocument (cmdlet)'           = { Convert-OfficeDocument -Text $markdown -From Markdown -To Html }
        'New-OfficeMarkdownHtmlOptions (cmdlet floor)' = { New-OfficeMarkdownHtmlOptions }
        'ConvertFrom-Markdown (cmdlet)'             = { ConvertFrom-Markdown -InputObject $markdown }
    }
    foreach ($lane in $lanes.GetEnumerator()) {
        [pscustomobject]@{
            Input        = $inputName
            InputChars   = $markdown.Length
            Lane         = $lane.Key
            MedianMicros = [Math]::Round((Measure-Median -Operation $lane.Value -Rounds $Rounds -Iterations $Iterations), 1)
        }
    }
    $stageMedians = [OfficeIMOMarkdownStages]::Measure($markdown, $Rounds, $Iterations * 4, $WarmupMilliseconds)
    $stageNames = [OfficeIMOMarkdownStages]::Names
    for ($index = 0; $index -lt $stageNames.Count; $index++) {
        [pscustomobject]@{
            Input        = $inputName
            InputChars   = $markdown.Length
            Lane         = '  engine: ' + $stageNames[$index]
            MedianMicros = [Math]::Round($stageMedians[$index], 1)
        }
    }
}

$firstCallResults = $null
if (-not $SkipFirstCall) {
    $pwsh = (Get-Process -Id $PID).Path
    $firstCallScript = @'
param($Manifest, $Markdown, $Tool)
$ErrorActionPreference = 'Stop'
$watch = [Diagnostics.Stopwatch]::StartNew()
if ($Tool -eq 'PSWriteOffice') { Import-Module $Manifest } else { Import-Module Microsoft.PowerShell.Utility }
$import = $watch.Elapsed.TotalMilliseconds
$watch.Restart()
$html = if ($Tool -eq 'PSWriteOffice') { [string](Convert-OfficeDocument -Text $Markdown -From Markdown -To Html) } else { (ConvertFrom-Markdown -InputObject $Markdown).Html }
$elapsed = $watch.Elapsed.TotalMilliseconds
if ($html.IndexOf('Typed document traversal', [StringComparison]::Ordinal) -lt 0) { throw "$Tool returned unexpected HTML on its first call." }
[string]::Format([Globalization.CultureInfo]::InvariantCulture, "{0}|{1}", $import, $elapsed)
'@
    $firstCallPath = Join-Path ([IO.Path]::GetTempPath()) ("officeimo-markdown-first-call-{0}.ps1" -f $PID)
    Set-Content -LiteralPath $firstCallPath -Value $firstCallScript
    try {
        $firstCallResults = @(foreach ($tool in 'PSWriteOffice', 'ConvertFrom-Markdown') {
            $samples = foreach ($run in 1..3) {
                $output = @(& $pwsh -NoLogo -NoProfile -File $firstCallPath -Manifest $moduleManifest -Markdown $inputs['SmallReadme'] -Tool $tool)
                if ($LASTEXITCODE -ne 0) { throw "The fresh-process $tool run failed: $($output -join [Environment]::NewLine)" }
                $parts = @(([string]$output[-1]) -split "\|")
                if ($parts.Count -ne 2) { throw "The fresh-process $tool run returned no timing record: $($output -join [Environment]::NewLine)" }
                , ($parts | ForEach-Object { [double]::Parse($_, [Globalization.CultureInfo]::InvariantCulture) })
            }
            $sortedImport = @($samples | ForEach-Object { $_[0] } | Sort-Object)
            $sortedCall = @($samples | ForEach-Object { $_[1] } | Sort-Object)
            [pscustomobject]@{ Tool = $tool; Input = 'SmallReadme'; ModuleImportMilliseconds = [Math]::Round($sortedImport[1], 1); FirstCallMilliseconds = [Math]::Round($sortedCall[1], 1) }
        })
    } finally {
        Remove-Item -LiteralPath $firstCallPath -ErrorAction SilentlyContinue
    }
}

if ($AsObject) { return [pscustomobject]@{ Warm = $results; FreshProcess = $firstCallResults } }

"PowerShell $($PSVersionTable.PSVersion); ConvertFrom-Markdown uses Markdig.Signed $markdigVersion; $Rounds rounds x $Iterations calls"
('Median microseconds per call after {0} ms warm-up per lane; input sizes: ' -f $WarmupMilliseconds) + (($Corpus | ForEach-Object { '{0} {1} chars' -f $_, $inputs[$_].Length }) -join ', ')
$rows = [ordered]@{}
foreach ($entry in $results) {
    if (-not $rows.Contains($entry.Lane)) { $rows[$entry.Lane] = [ordered]@{ Lane = $entry.Lane } }
    $rows[$entry.Lane][$entry.Input] = $entry.MedianMicros
}
$rows.Values | ForEach-Object { [pscustomobject]$_ } | Format-Table -Property (@('Lane') + @($Corpus)) -AutoSize | Out-String -Width 200
if ($firstCallResults) {
    'Fresh process, SmallReadme input (median of three processes):'
    $firstCallResults | Format-Table Tool, ModuleImportMilliseconds, FirstCallMilliseconds -AutoSize | Out-String -Width 200
}
