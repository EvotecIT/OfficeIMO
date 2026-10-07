param([switch]$IncludeLinux)
$ErrorActionPreference = 'Stop'
$comparisons = @()
$observations = 0
$measurements = 0
$sources = @()
$files = @('retained-native-current.json', 'retained-native-current-macos.json')
if ($IncludeLinux) { $files += 'retained-native-current-linux.json' }
foreach ($file in $files) {
    $packet = Get-Content -LiteralPath (Join-Path $PSScriptRoot $file) -Raw | ConvertFrom-Json
    $hostLabel = if ($file -like '*macos*') { 'macOS' } elseif ($file -like '*linux*') { 'Linux (WSL2)' } else { 'Windows' }
    $sources += [pscustomobject]@{ Host = $hostLabel; Head = $packet.Head; Status = $packet.Status; Source = $packet.Source; Sha256 = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot $file)).Hash }
    foreach ($run in $packet.Reports) {
        $observations += $run.Report.Benchmarks.Count
        foreach ($benchmark in $run.Report.Benchmarks) {
            if ($benchmark.Statistics.N -ne 12 -or $benchmark.Memory.TotalOperations -ne 4) { throw 'Incomplete CSV native observations' }
            $measurements += $benchmark.Statistics.N
        }
        foreach ($group in $run.Report.Benchmarks | Group-Object { $_.Type + ':' + $_.Parameters }) {
            $peerMethod = if ($group.Group[0].Type -eq 'CsvFileWriteBenchmarks') { 'CsvHelper' } else { 'Sylvan_WriteDataReader' }
            $peers = @($group.Group | Where-Object Method -eq $peerMethod)
            if ($peers.Count -ne 1) { throw 'CSV peer join is ambiguous' }
            $peer = $peers[0]
            foreach ($office in $group.Group | Where-Object { $_.Method -like 'OfficeIMO*' }) {
                $comparisons += [pscustomobject]@{
                    Host = $hostLabel
                    Runtime = $run.Runtime
                    Placement = $run.Mask
                    Workload = $office.Type
                    Case = $office.Parameters
                    Method = $office.Method
                    Peer = $peer.Method
                    OfficeMedianMs = $office.Statistics.Median / 1e6
                    PeerMedianMs = $peer.Statistics.Median / 1e6
                    MedianRatio = $office.Statistics.Median / $peer.Statistics.Median
                    OfficeAllocatedBytes = $office.Memory.BytesAllocatedPerOperation
                    PeerAllocatedBytes = $peer.Memory.BytesAllocatedPerOperation
                    AllocationRatio = $office.Memory.BytesAllocatedPerOperation / $peer.Memory.BytesAllocatedPerOperation
                    SamplesPerEngine = 12
                }
            }
        }
    }
}
$expectedObservations = if ($IncludeLinux) { 368 } else { 276 }
$expectedMeasurements = if ($IncludeLinux) { 4416 } else { 3312 }
$expectedComparisons = if ($IncludeLinux) { 208 } else { 156 }
if ($observations -ne $expectedObservations -or $measurements -ne $expectedMeasurements -or $comparisons.Count -ne $expectedComparisons) { throw 'Wrong current CSV matrix inventory' }
$reference = @{}
foreach ($source in $sources[0].Source) { $reference[$source.Path] = $source.Sha256 }
foreach ($packet in $sources | Select-Object -Skip 1) {
    if ($sources[0].Source.Count -ne $packet.Source.Count) { throw 'Cross-host CSV source inventory differs' }
    foreach ($source in $packet.Source) { if ($reference[$source.Path] -ne $source.Sha256) { throw "Cross-host CSV source differs: $($source.Path)" } }
}
$summaryName = if ($IncludeLinux) { 'summary-current-linux.json' } else { 'summary-current.json' }
[ordered]@{
    Observations = $observations
    RetainedMeasurements = $measurements
    Comparisons = $comparisons
    Sources = $sources
    Policy = 'Native .NET8/.NET10; 24 warmups; 12 measurements; four invocations; no outlier removal; output validation outside timing; file writes include creation/close with equal buffers and bytes; OS flush, not durable storage; mixed typed inputs validate every decoded field'
} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $PSScriptRoot $summaryName) -Encoding utf8
$comparisons | Group-Object Host,Runtime,Workload | ForEach-Object {
    [pscustomobject]@{ Group = $_.Name; Comparisons = $_.Count; SlowerMedians = @($_.Group | Where-Object MedianRatio -gt 1).Count; LowerAllocation = @($_.Group | Where-Object AllocationRatio -lt 1).Count; MedianRatioMin = ($_.Group.MedianRatio | Measure-Object -Minimum).Minimum; MedianRatioMax = ($_.Group.MedianRatio | Measure-Object -Maximum).Maximum }
} | Format-Table -AutoSize
