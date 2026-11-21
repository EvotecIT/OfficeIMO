$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.Word.Tests/bin/Release/net8.0')
$parentCounts = Get-BenchmarkInput Parents 0, 32, 128 -Int
if ([Environment]::Version.Major -lt 8) { throw 'Run this suite with PowerShell on .NET 8 or newer.' }
$references = @([IO.Directory]::GetFiles((Join-Path $PSHOME 'ref'), '*.dll'))
foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.Word', 'DocumentFormat.OpenXml', 'DocumentFormat.OpenXml.Framework', 'System.IO.Packaging') {
    $path = Join-Path $binaryRoot "$name.dll"
    [void] [Reflection.Assembly]::LoadFrom($path)
    $references += $path
}
# Newer PowerShell runtimes unify the net8.0 library's System.Runtime reference.
Add-Type -Path (Join-Path $PSScriptRoot 'WordSharedCloneWorkload.cs') -ReferencedAssemblies $references -CompilerOptions '/nowarn:1701'

New-BenchmarkSuite 'officeimo-word-shared-clone' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 7 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Load, public legacy-write assessment and disposal; one 64 KiB inert payload shared by root/main and 0, 32 or 128 additional parents. Setup and input-graph validation are outside measurement. Every operation disposes its document, including warmups and failures.'
    Add-BenchmarkMetadata WordSha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.Word.dll') -Algorithm SHA256).Hash
    Add-BenchmarkCaseSource {
        foreach ($parents in $parentCounts) { [pscustomobject]@{ Name = "Parents-$parents"; Parents = $parents } }
    }
    Set-BenchmarkSetup { param($case, $run) $run.Workload = [WordSharedCloneWorkload]::new($case.Parents) }
    Add-BenchmarkEngine OfficeIMO {
        Add-BenchmarkOperation Assess { param($case, $run) $run.Workload.Assess() }
    }
    Add-BenchmarkValidation { param($case, $run) $run.Workload.Validate() }
    Add-BenchmarkMetric OperationAllocatedBytes { param($case, $run) $run.Workload.AllocatedBytes }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
