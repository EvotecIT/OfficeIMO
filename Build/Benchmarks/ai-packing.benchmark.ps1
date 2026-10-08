$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.AI/bin/Release/net10.0')
$blockCounts = Get-BenchmarkInput Blocks 100, 1000, 4000 -Int
$requestBudgets = Get-BenchmarkInput RequestCharacters 48000, 2000000 -Int
if ([Environment]::Version.Major -lt 10) { throw 'Run this suite with PowerShell on .NET 10 or newer.' }
$references = @([IO.Directory]::GetFiles((Join-Path $PSHOME 'ref'), '*.dll'))
foreach ($name in 'OfficeIMO.Reader.Core', 'OfficeIMO.AI') {
    $assemblyPath = Join-Path $binaryRoot "$name.dll"
    [void] [Reflection.Assembly]::LoadFrom($assemblyPath)
    $references += $assemblyPath
}
Add-Type -Path (Join-Path $PSScriptRoot 'AiPackingWorkload.cs') -ReferencedAssemblies $references

New-BenchmarkSuite 'officeimo-ai-packing' -OutputRoot (Join-Path $repositoryRoot '.validation/ai-packing') {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 7 -Order Rotated -OutlierMode None
    Add-BenchmarkMetadata Contract 'Identical captured short text blocks; request planning and validation, excluding source capture and model inference.'
    Add-BenchmarkMetadata BinarySha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.AI.dll') -Algorithm SHA256).Hash
    Add-BenchmarkCaseSource {
        foreach ($blocks in $blockCounts) {
            foreach ($characters in $requestBudgets) {
                [pscustomobject]@{ Name = "Blocks-$blocks-Budget-$characters"; Blocks = $blocks; RequestCharacters = $characters }
            }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = [AiPackingWorkload]::new($case.Blocks, $case.RequestCharacters)
    }
    Add-BenchmarkEngine OfficeIMO {
        Add-BenchmarkOperation Pack { param($case, $run) $run.Workload.Run() }
    }
    Add-BenchmarkValidation { param($case, $run) $run.Workload.Validate() }
    Add-BenchmarkMetric OperationAllocatedBytes { param($case, $run) $run.Workload.AllocatedBytes }
    Add-BenchmarkMetric MeasurementCalls { param($case, $run) $run.Workload.MeasurementCalls }
    Add-BenchmarkMetric RequestCount { param($case, $run) $run.Workload.RequestCount }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
