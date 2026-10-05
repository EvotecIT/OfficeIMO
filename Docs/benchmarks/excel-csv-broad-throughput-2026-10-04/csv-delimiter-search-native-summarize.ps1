param([string[]]$Files=@('retained-native-csv-delimiter-scan-v1-65535.json','retained-native-csv-delimiter-scan-v1-4294901760.json'),[string]$Output='summary-native.json')
$ErrorActionPreference='Stop'
$pairs=@();$sources=@();$observations=0;$measurements=0
$expectedCases=@((Get-Content -LiteralPath (Join-Path $PSScriptRoot 'cases.json') -Raw | ConvertFrom-Json).PSObject.Properties).Count
foreach ($file in $Files) {
    $path=Join-Path $PSScriptRoot $file
    $packet=Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
    $hostLabel=if($file -like '*macos*'){'macOS'}else{'Windows'}
    foreach($run in $packet.Reports) {
        if($run.Report.Benchmarks.Count -ne 2*$expectedCases){throw 'Wrong CSV snapshot observation count'}
        foreach($benchmark in $run.Report.Benchmarks){
            if($benchmark.Statistics.N -ne 12 -or $benchmark.Memory.TotalOperations -ne 4){throw 'Incomplete CSV measurement policy'}
            $observations++;$measurements+=$benchmark.Statistics.N
        }
        $sources+=[pscustomobject]@{Host=$hostLabel;Runtime=$run.Runtime;Placement=$packet.Mask;BeforeSource=$run.Manifest.BeforeSource;AfterSource=$run.Manifest.AfterSource;RawSha256=(Get-FileHash -LiteralPath $path).Hash}
        foreach($group in $run.Report.Benchmarks | Group-Object {$_.FullName.Replace('SnapshotComparison.Before(', 'SnapshotComparison.Pair(').Replace('SnapshotComparison.After(', 'SnapshotComparison.Pair(')}){
            $before=@($group.Group | Where-Object Method -eq 'Before');$after=@($group.Group | Where-Object Method -eq 'After')
            if($before.Count -ne 1 -or $after.Count -ne 1){throw 'Ambiguous complete CSV scenario join'}
            $pairs+=[pscustomobject]@{Host=$hostLabel;Runtime=$run.Runtime;Placement=$packet.Mask;Case=$group.Name;BeforeMedianMs=$before[0].Statistics.Median/1e6;AfterMedianMs=$after[0].Statistics.Median/1e6;MedianRatio=$after[0].Statistics.Median/$before[0].Statistics.Median;BeforeAllocatedBytes=$before[0].Memory.BytesAllocatedPerOperation;AfterAllocatedBytes=$after[0].Memory.BytesAllocatedPerOperation;AllocationRatio=$after[0].Memory.BytesAllocatedPerOperation/$before[0].Memory.BytesAllocatedPerOperation}
        }
    }
}
$reference=@{}
foreach($file in $sources[0].AfterSource.Source){$reference[$file.Path]=$file.Sha256}
foreach($source in $sources | Select-Object -Skip 1){
    if($source.AfterSource.Source.Count -ne $reference.Count){throw 'CSV source inventory differs'}
    foreach($file in $source.AfterSource.Source){if($reference[$file.Path] -ne $file.Sha256){throw "CSV candidate source differs: $($file.Path)"}}
}
if($pairs.Count*2 -ne $observations){throw 'Missing CSV comparison pairs'}
[ordered]@{Observations=$observations;RetainedMeasurements=$measurements;Comparisons=$pairs;Sources=$sources;Policy='Native actual .NET8/.NET10; 16 warmups, twelve retained samples, four complete validated writes, no outlier removal; same harness/dependencies; complete byte/field validation; full scenario names used for joins'}|ConvertTo-Json -Depth 14|Set-Content -LiteralPath (Join-Path $PSScriptRoot $Output) -Encoding utf8
$pairs | Group-Object Host,Runtime,Placement | ForEach-Object {[pscustomobject]@{Group=$_.Name;Comparisons=$_.Count;SlowerMedians=@($_.Group | Where-Object MedianRatio -gt 1).Count;IncreasedAllocation=@($_.Group | Where-Object AllocationRatio -gt 1).Count}}
