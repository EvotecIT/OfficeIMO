param(
 [Parameter(Mandatory)][string]$SnapshotRoot,
 [Parameter(Mandatory)][string]$CompareDirectory,
 [Parameter(Mandatory)][string]$ModulePath,
 [long]$Mask=65535
)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$module=Import-Module -Name $ModulePath -PassThru
$moduleRoot=Split-Path $ModulePath -Parent
$env:OFFICEIMO_PERFORMANCE_SNAPSHOTS=Join-Path $SnapshotRoot 'net10.0'
$env:OFFICEIMO_COMPARISON_CASES=Join-Path $root 'cases.json'
$env:OFFICEIMO_PERFORMANCE_AA='0'
$env:OFFICEIMO_BENCHMARK_DATA=Join-Path $root 'fixtures'
$env:OFFICEIMO_BENCHMARK_OUTPUT=Join-Path $root 'fixtures'
if($IsWindows -or $IsLinux){[Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity=[IntPtr]::new($Mask)}
[Diagnostics.Process]::GetCurrentProcess().PriorityClass='Normal'
$manifest=Get-Content -LiteralPath (Join-Path $SnapshotRoot 'manifest-net10.0.json') -Raw|ConvertFrom-Json
function Assert-Frozen {
 foreach($file in $manifest.Manifest){if((Get-FileHash -LiteralPath (Join-Path $SnapshotRoot "net10.0/candidate-excel/$($file.Name)")).Hash -ne $file.After){throw 'Current public-reader snapshot changed'}}
}
Assert-Frozen
[Reflection.Assembly]::LoadFrom((Join-Path $CompareDirectory 'bin/Release/net10.0/Compare.dll'))|Out-Null
$global:CurrentPublicProbes=@{};$global:CurrentPublicExpected=@{}
$fixtures=[ordered]@{}
$flags=[Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
foreach($rows in 1000,25000,250000,1000000){
 $scenario="Rows-$rows";$probes=@{};$observations=@{};$hashes=@{}
 foreach($engine in 'OfficeIMO','Sylvan','ExcelReaderNet'){
  $probe=[SnapshotProbe]::Load("Public-$rows-$engine",'candidate')
  $probes[$engine]=$probe;$observations[$engine]=$probe.RunSynchronously().ToString()
  $instance=[SnapshotProbe].GetField('_instance',$flags).GetValue($probe)
  $path=[string]$instance.GetType().GetField('_path',$flags).GetValue($instance)
  $zip=[IO.Compression.ZipFile]::OpenRead($path);$parts=[ordered]@{}
  try{foreach($entry in $zip.Entries|Where-Object {$_.FullName -match '^xl/(worksheets/|styles.xml$|sharedStrings.xml$)'}|Sort-Object FullName){
   $stream=$entry.Open();$hash=[Security.Cryptography.SHA256]::Create()
   try{$parts[$entry.FullName]=[Convert]::ToHexString($hash.ComputeHash($stream))}finally{$hash.Dispose();$stream.Dispose()}
  }}finally{$zip.Dispose()}
  $hashes[$engine]=$parts
 }
 foreach($engine in 'Sylvan','ExcelReaderNet'){
  [SnapshotProbe]::ValidateEquivalent($probes.OfficeIMO,$probes[$engine])
  if($hashes[$engine].Count -ne $hashes.OfficeIMO.Count){throw 'Fixture part inventory differs'}
  foreach($part in $hashes.OfficeIMO.Keys){if($hashes.OfficeIMO[$part] -ne $hashes[$engine][$part]){throw 'Decoded peer input differs'}}
 }
 $global:CurrentPublicProbes[$scenario]=$probes;$global:CurrentPublicExpected[$scenario]=$observations
 $fixtures[$scenario]=[ordered]@{DecodedParts=$hashes;Observations=$observations;CompleteFieldValidation='Existing engine-specific benchmark setup verifies every header, value, type and row; observations equal'}
}
$fixtures|ConvertTo-Json -Depth 9|Set-Content -LiteralPath (Join-Path $root "fixtures-$Mask.json") -Encoding utf8
$output=Join-Path $root "rotated-$Mask"
if(Test-Path -LiteralPath $output){throw 'Refuse existing current reader measurements'}
try{
 Invoke-BenchmarkSuite -OutputRoot $output -WarmupCount 16 -IterationCount 24 -RunOrder Rotated -OutlierMode None -Settings {
  New-BenchmarkSuite 'Complete public typed XLSX reads' {
   Add-BenchmarkMetadata AffinityMask $(if($IsWindows -or $IsLinux){$Mask}else{'OS scheduled'})
   Add-BenchmarkMetadata Runtime ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
   Add-BenchmarkMetadata Source 'Qualified worksheet-encoding V2 source; identical frozen source/dependencies for every engine; all selected fields of every row consumed'
   Add-BenchmarkMetadata Normalization 'One complete public typed read per sample; producers/validation outside timing'
   Add-BenchmarkMetadata SylvanVersion '0.5.8'
   Add-BenchmarkMetadata ExcelReaderNetVersion '5.1.1'
   Add-BenchmarkMetadata Licenses 'Sylvan.Data.Excel 0.5.8 embedded license.txt; ExcelReader.NET 5.1.1 MIT nuspec; benchmark-only dependencies'
   Add-BenchmarkMetadata ControlModuleVersion ($module.Version.ToString())
   Add-BenchmarkMetadata ControlOwnerSha256 ((Get-FileHash -LiteralPath (Join-Path $moduleRoot 'Lib/Core/PowerForge.dll')).Hash)
   Add-BenchmarkMetadata ControlCmdletsSha256 ((Get-FileHash -LiteralPath (Join-Path $moduleRoot 'Lib/Core/PSPublishModule.dll')).Hash)
   Add-BenchmarkCaseSource @($global:CurrentPublicProbes.Keys|Sort-Object|ForEach-Object {[pscustomobject]@{Name=$_}})
   Set-BenchmarkSetup {param($case,$run)
    $run.Probe=$global:CurrentPublicProbes[$case.Scenario][$case.Engine]
    $run.Expected=$global:CurrentPublicExpected[$case.Scenario][$case.Engine]
   }
   foreach($engine in 'OfficeIMO','Sylvan','ExcelReaderNet'){
    Add-BenchmarkEngine $engine {Add-BenchmarkOperation ReadAll {param($case,$run) $run.Result=$run.Probe.RunSynchronously()}}
   }
   Add-BenchmarkValidation {param($case,$run) if($run.Result.ToString() -ne $run.Expected){throw 'Complete typed read differs from independently validated output'}}
   Add-BenchmarkComparison -Baseline OfficeIMO -Metric MeanMs,MedianMs -TieTolerance 0.05
  }
 }
}finally{foreach($set in $global:CurrentPublicProbes.Values){foreach($probe in $set.Values){$probe.Dispose()}}}
Assert-Frozen
$files=@(Get-ChildItem -LiteralPath $output -Filter run-report.json -Recurse)
if($files.Count -ne 1){throw 'Wrong current public-reader report inventory'}
$report=Get-Content -LiteralPath $files[0].FullName -Raw|ConvertFrom-Json
if($report.summary.Count -ne 12 -or @($report.summary|Where-Object {$_.sampleCount -ne 24 -or $_.failureCount -ne 0}).Count){throw 'Current public-reader comparison incomplete'}
[ordered]@{Manifest=$manifest;FixtureQualification=$fixtures;Report=$report;Sha256=(Get-FileHash -LiteralPath $files[0].FullName).Hash;RunnerSha256=(Get-FileHash -LiteralPath $PSCommandPath).Hash;Policy='16 warmups;24 retained samples;rotated engine order;one complete read per sample;no outlier removal;current frozen OfficeIMO source;all selected fields and decoded fixture bytes qualified'}|ConvertTo-Json -Depth 26|Set-Content -LiteralPath (Join-Path $root "retained-current-$Mask.json") -Encoding utf8
Write-Output "Current public reader comparison qualified:12 observations/$Mask"
