$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.Migration.Benchmarks/bin/Release/net10.0')
$profiles = (Get-BenchmarkInput Profiles 'Word69,Biff4,IWork09,Dbx,OfficeXml') -split ','
$scales = (Get-BenchmarkInput Scales 'Small,Medium,Large') -split ','
$measureRetainedMemory = Get-BenchmarkInput MeasureRetainedMemory $false -Bool
$productSource = Get-BenchmarkInput ProductSource ''
$workloads = @{}
[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.Migration.Benchmarks.dll'))

New-BenchmarkSuite 'officeimo-migration' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 5 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Contract 'Matched deterministic inputs supported by baseline and candidate. LoadInspect includes full public load and inspection/read materialization. ConvertSave includes load plus ordinary owner conversion/save. Complete ordered semantics, diagnostics and saved artifact/schema validation occur outside the measured operation. DBX SearchContent separately measures eager content search. Synthetic scale inputs are not independent-producer compatibility evidence.'
    Add-BenchmarkMetadata Scales 'Word69 shared-FKP runs and OfficeXml mapped-property paragraphs: 100/1000/10000; Biff4 finding-heavy rows and XF declarations: 1000/4000/16000; IWork09 shared-image drawables and Dbx indexed messages: 6/64/512.'
    Add-BenchmarkMetadata Diagnostics 'Every supported result is fully validated. BIFF4 retains all per-cell finding locations followed by one global number-format finding; known loss is explicitly accepted for XLSX conversion. Legacy iWork09 explicitly permits the existing content-limited diagnostic under partial editable policy while validating every selected image; the other common workloads require their exact declared diagnostic contract.'
    Add-BenchmarkMetadata Measurement 'PowerForge operation elapsed time and managed allocation include the invocation boundary. Input generation, expected values and result validation are outside timing. Retained managed observations and sampled process maxima are separate opt-in cohorts; sampled maxima are observed lower bounds and include the host/observer.'
    Add-BenchmarkMetadata Placement 'One applied affinity domain per cohort. An omitted affinity mask inherits placement and does not qualify per-domain behavior. Use identical masks, policy and isolated fresh processes for all revisions.'
    if ($productSource) { Add-BenchmarkMetadata ProductSource $productSource }
    Add-BenchmarkMetadata Runtime ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    Add-BenchmarkMetadata RunnerModuleSha256 (Get-FileHash (Get-Command Invoke-BenchmarkSuite).ImplementingType.Assembly.Location -Algorithm SHA256).Hash
    Add-BenchmarkMetadata WorkloadSha256 (Get-FileHash (Join-Path $binaryRoot 'OfficeIMO.Migration.Benchmarks.dll') -Algorithm SHA256).Hash
    Add-BenchmarkMetadata SpecSha256 (Get-FileHash (Join-Path $PSScriptRoot 'migration.benchmark.ps1') -Algorithm SHA256).Hash
    foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.Word', 'OfficeIMO.Excel', 'OfficeIMO.Email', 'OfficeIMO.Rtf', 'OfficeIMO.IWork', 'OfficeIMO.Word.IWork') {
        Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash (Join-Path $binaryRoot "$name.dll") -Algorithm SHA256).Hash
    }
    Add-BenchmarkCaseSource {
        foreach ($family in $profiles) {
            foreach ($size in $scales) {
                $index = @('Small', 'Medium', 'Large').IndexOf($size)
                $units = switch ($family) {
                    Biff4 { @(1000, 4000, 16000)[$index] }
                    { $_ -in 'IWork09', 'Dbx' } { @(6, 64, 512)[$index] }
                    default { @(100, 1000, 10000)[$index] }
                }
                $key = "$family-$size"
                $workloads[$key] = [OfficeIMO.Migration.Benchmarks.MigrationWorkload]::Create($family, $units)
                [pscustomobject]@{ Name = $key; Family = $family; Scale = $size; Units = $units; InputSha256 = $workloads[$key].InputSha256 }
            }
        }
    }
    Add-BenchmarkSkipRule {
        param($case, $run)
        if ($case.Operation -eq 'SearchContent' -and $case.Family -ne 'Dbx') { 'Content search belongs to the DBX store contract.' }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = $workloads["$($case.Family)-$($case.Scale)"]
        $run.Workload.ReleaseResults()
        if ($measureRetainedMemory) { $run.MemoryProbe = [PowerForge.BenchmarkManagedMemoryProbe]::new() }
    }
    Add-BenchmarkEngine OfficeIMO {
        foreach ($operation in 'LoadInspect', 'ConvertSave', 'SearchContent') {
            Add-BenchmarkOperation $operation {
                param($case, $run)
                $run.Workload.Execute($case.Operation)
            }
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $run.Workload.Validate()
        Assert-BenchmarkValue -Actual $run.Workload.VerifiedUnits -Expected $case.Units
        $run.OutputBytes = $run.Workload.OutputBytes
        $run.Workload.ReleaseResults()
        if ($measureRetainedMemory) { $run.RetainedMemory = $run.MemoryProbe.Capture() }
    }
    Add-BenchmarkMetric VerifiedUnits { param($case, $run) $run.Workload.VerifiedUnits }
    Add-BenchmarkMetric InputBytes { param($case, $run) $run.Workload.InputBytes }
    Add-BenchmarkMetric OutputBytes { param($case, $run) $run.OutputBytes }
    if ($measureRetainedMemory) {
        Add-BenchmarkMetric ManagedBaselineBytes { param($case, $run) $run.RetainedMemory.BaselineBytes }
        Add-BenchmarkMetric CollectedManagedBytes { param($case, $run) $run.RetainedMemory.CollectedBytes }
        Add-BenchmarkMetric RetainedManagedDeltaBytes { param($case, $run) $run.RetainedMemory.DeltaBytes }
    }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
