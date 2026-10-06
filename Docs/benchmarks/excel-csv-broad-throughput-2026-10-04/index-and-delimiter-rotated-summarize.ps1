param([Parameter(Mandatory)][string[]]$Files,[Parameter(Mandatory)][string]$Output,[switch]$Reversed)
$ErrorActionPreference='Stop'
$pairs=@();$sources=@();$observations=0;$retained=0
foreach($file in $Files){
 $packet=Get-Content -LiteralPath $file -Raw|ConvertFrom-Json
 $report=$packet.Report
 $metadata=$report.metadata
 $expectedSamples=[int]$metadata.iterationCount
 $operations=[int]$metadata.'benchmark.OperationsPerSample'
 if($operations -le 0 -or $expectedSamples -le 0 -or $metadata.outlierMode -ne 'None' -or $metadata.runOrder -ne 'Rotated'){throw 'Unexpected rotated measurement policy'}
 $hostLabel=if($metadata.osLabel -eq 'macOS' -or $file -like '*macos*'){'macOS'}else{'Windows'}
 foreach($group in $report.summary|Group-Object scenario){
  $engines=@{}
  foreach($engine in 'Before','Control','After'){
   $row=@($group.Group|Where-Object engine -eq $engine)
   if($row.Count -ne 1 -or $row[0].sampleCount -ne $expectedSamples -or $row[0].failureCount -ne 0 -or $row[0].status -ne 'Succeeded'){throw 'Missing or failed rotated observation'}
   $samples=@($report.samples|Where-Object{$_.scenario -eq $group.Name -and $_.engine -eq $engine})
   if($samples.Count -ne $expectedSamples -or @($samples|Where-Object status -ne 'Succeeded').Count){throw 'Retained rotated sample inventory differs'}
   $engines[$engine]=$row[0]
   $observations++;$retained+=$expectedSamples
  }
  $before=$engines.Before.medianMs/$operations;$after=$engines.After.medianMs/$operations;$control=$engines.Control.medianMs/$operations
  $pairs+=[pscustomobject]@{Host=$hostLabel;Runtime=$metadata.'benchmark.Runtime';Case=$group.Name;BeforeMedianMs=$before;AfterMedianMs=$after;ControlMedianMs=$control;AfterBeforeRatio=$after/$before;ControlBeforeRatio=$control/$before;CandidateOverBaseline=$(if($Reversed){$before/$after}else{$after/$before});CandidateCopyRatio=$(if($Reversed){$control/$before}else{$control/$before});RunId=$report.runId}
 }
 $sources+=[pscustomobject]@{File=[IO.Path]::GetFileName($file);Sha256=(Get-FileHash -LiteralPath $file).Hash;Host=$hostLabel;Runtime=$metadata.'benchmark.Runtime';Pwsh=$metadata.pwsh;RunId=$report.runId;StartedUtc=$report.startedUtc;FinishedUtc=$report.finishedUtc;SourceHead=$metadata.gitSha;SourceBranch=$metadata.gitBranch;WorktreeClean=$metadata.gitWorktreeClean;ControlOwnerSha256=$metadata.'benchmark.ControlOwnerSha256';ControlCmdletsSha256=$metadata.'benchmark.ControlCmdletsSha256';Warmup=$metadata.warmupCount;Iterations=$expectedSamples;OperationsPerSample=$operations;Diagnostic=$metadata.'benchmark.Diagnostic';Policy=$packet.Policy}
}
[ordered]@{Observations=$observations;RetainedMeasurements=$retained;Reversed=$Reversed.IsPresent;Policy='All retained samples; three engines rotated within each iteration; per-operation medians; no outlier removal. Forward Before/Control use identical baseline bytes. Reverse Before/Control use identical candidate bytes and After uses baseline.';Comparisons=$pairs;Sources=$sources}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $Output -Encoding utf8
Write-Output "Qualified rotated summary: $observations observations, $retained retained measurements"
