param([Parameter(Mandatory)][string]$Workspace)
$ErrorActionPreference='Stop'
Set-Location $Workspace
$root=$PSScriptRoot
$source=@(foreach($directory in 'OfficeIMO.CSV','OfficeIMO.Core','OfficeIMO.SharedSource','OfficeIMO.CSV.Benchmarks','Benchmarks'){
 Get-ChildItem -LiteralPath $directory -File -Recurse|Where-Object {$_.FullName -notmatch '[\\/](bin|obj)[\\/]' -and $_.Extension -in '.cs','.csproj'}|ForEach-Object {
  $bytes=[Text.Encoding]::UTF8.GetBytes([IO.File]::ReadAllText($_.FullName).Replace("`r`n","`n"))
  [pscustomobject]@{Path=[IO.Path]::GetRelativePath($Workspace,$_.FullName).Replace('\','/');Sha256=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))}
 }
})
$reports=@()
$masks=if($IsWindows){@([long]65535,[long]4294901760)}else{@([long]65535)}
$env:OFFICEIMO_BENCHMARK_OUTPUT=Join-Path $root 'fixtures'
$env:OFFICEIMO_BENCHMARK_DATA=Join-Path $root 'fixtures'
foreach($tfm in 'net10.0','net8.0'){
 & dotnet build OfficeIMO.CSV.Benchmarks/OfficeIMO.CSV.Benchmarks.csproj -c Release -f $tfm *> (Join-Path $root "build-$tfm.log")
 if($LASTEXITCODE){throw "CSV benchmark build failed $tfm"}
}
foreach($tfm in 'net10.0','net8.0'){
 $binaryRoot=Join-Path $Workspace "OfficeIMO.CSV.Benchmarks/bin/Release/$tfm"
 $binaries=@(Get-ChildItem -LiteralPath $binaryRoot -File|ForEach-Object {[pscustomobject]@{Name=$_.Name;Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash}})
 foreach($mask in $masks){
  $output=Join-Path $root "native-$tfm-$mask"
  if(Test-Path -LiteralPath $output){throw 'Refuse existing CSV observations'}
  $arguments=@((Join-Path $binaryRoot 'OfficeIMO.CSV.Benchmarks.dll'),'--filter','*CsvFileWriteBenchmarks*','*CsvDataReaderWriteBenchmarks*','--warmupCount','24','--iterationCount','12','--invocationCount','4','--unrollFactor','1','--outliers','DontRemove','--artifacts',$output,'--priority','Normal')
  if($IsWindows){$arguments+=@('--affinityMasks',[string]$mask)}
  & dotnet @arguments *> (Join-Path $root "native-$tfm-$mask.log")
  if($LASTEXITCODE){throw "CSV native job failed $tfm/$mask"}
  $measurements=@(Get-ChildItem -LiteralPath (Join-Path $output 'results') -Filter '*-report-full.json')
  if($measurements.Count -ne 2){throw 'Wrong CSV report inventory'}
  foreach($file in $measurements){
   $report=Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json
   $expected=if($file.Name -match 'CsvFileWriteBenchmarks'){28}else{18}
   if($report.Benchmarks.Count -ne $expected -or @($report.Benchmarks|Where-Object {$_.Statistics.N -ne 12 -or $_.Memory.TotalOperations -le 0}).Count){throw 'CSV benchmark report incomplete'}
   $reports+=[pscustomobject]@{Runtime=$tfm;Mask=$(if($IsWindows){$mask}else{'OS scheduled'});Report=$report;Sha256=(Get-FileHash -LiteralPath $file.FullName).Hash;Binaries=$binaries}
  }
  Write-Output "CSV file/mixed native job qualified $tfm/${mask}: 46 observations"
 }
 foreach($file in $binaries){if((Get-FileHash -LiteralPath (Join-Path $binaryRoot $file.Name)).Hash -ne $file.Sha256){throw 'CSV benchmark binary changed'}}
}
foreach($file in $source){
 $bytes=[Text.Encoding]::UTF8.GetBytes([IO.File]::ReadAllText((Join-Path $Workspace $file.Path)).Replace("`r`n","`n"))
 if([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes)) -ne $file.Sha256){throw 'CSV source changed during qualification'}
}
[ordered]@{Head=(& git rev-parse HEAD);Status=@(& git status --porcelain);Source=$source;Reports=$reports;Policy='Native actual .NET8/.NET10; 24 warmups; 12 retained measurements; four invocations; no outlier removal. File creation/close with identical buffers and output byte/field validation; mixed DataReader text exports validate each decoded field with nullable typed inputs. File writes flush to the OS, not durable media.'}|ConvertTo-Json -Depth 26|Set-Content -LiteralPath (Join-Path $root 'retained-native-current.json') -Encoding utf8
