param([Parameter(Mandatory)][string]$CompareDirectory,[long]$Mask=65535,[string]$Cases='cases.json',[string[]]$Runtimes=@('net10.0','net8.0'),[int]$Warmup=16,[int]$Iterations=12,[int]$Invocations=1)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$reports=@()
$casePath=Join-Path $root $Cases
$caseSpecs=Get-Content -LiteralPath $casePath -Raw|ConvertFrom-Json
$expectedCount=@($caseSpecs.PSObject.Properties).Count*2
foreach($tfm in $Runtimes){
 $manifest=Get-Content -LiteralPath (Join-Path $root "manifest-$tfm.json") -Raw|ConvertFrom-Json
 foreach($file in $manifest.Manifest){foreach($side in 'Before','After'){
  $directory=if($side -eq 'Before'){'baseline-excel'}else{'candidate-excel'}
  if((Get-FileHash -LiteralPath (Join-Path $root "$tfm/$directory/$($file.Name)")).Hash -ne $file.$side){throw 'Dimensionless worksheet snapshot changed'}
 }}
 $env:OFFICEIMO_PERFORMANCE_SNAPSHOTS=Join-Path $root $tfm
 $env:OFFICEIMO_COMPARISON_CASES=$casePath
 $env:OFFICEIMO_BENCHMARK_OUTPUT=Join-Path $root 'fixtures'
 $env:OFFICEIMO_BENCHMARK_DATA=Join-Path $root 'fixtures'
 $env:OFFICEIMO_COMPARISON_MASK=[string]$Mask
 $env:OFFICEIMO_PERFORMANCE_AA='0'
 Push-Location $CompareDirectory
 try{
  & dotnet "bin/Release/$tfm/Compare.dll" --validate *> (Join-Path $root "validate-$tfm.log")
  if($LASTEXITCODE){throw 'Dimensionless worksheet output validation failed'}
  $output=Join-Path $root "native-dimensionless-v2-$tfm-$Mask"
  if(Test-Path -LiteralPath $output){throw 'Refuse existing dimensionless measurements'}
  & dotnet "bin/Release/$tfm/Compare.dll" --filter '*SnapshotComparison*' --warmupCount $Warmup --iterationCount $Iterations --invocationCount $Invocations --unrollFactor 1 --outliers DontRemove --artifacts $output *> (Join-Path $root "native-dimensionless-v2-$tfm-$Mask.log")
  if($LASTEXITCODE){throw 'Dimensionless worksheet native job failed'}
 }finally{Pop-Location}
 $path=Join-Path $output 'results/SnapshotComparison-report-full.json'
 $report=Get-Content -LiteralPath $path -Raw|ConvertFrom-Json
 if($report.Benchmarks.Count -ne $expectedCount -or @($report.Benchmarks|Where-Object {$_.Statistics.N -ne $Iterations -or $_.Memory.TotalOperations -le 0}).Count){throw 'Dimensionless worksheet report incomplete'}
 $reports+=[pscustomobject]@{Runtime=$tfm;Report=$report;Sha256=(Get-FileHash -LiteralPath $path).Hash;Manifest=$manifest;ValidationSha256=(Get-FileHash -LiteralPath (Join-Path $root "validate-$tfm.log")).Hash}
 Write-Output "Dimensionless worksheet native job qualified: $tfm / $expectedCount observations"
}
[ordered]@{Policy="Corrected dimensionless UTF8 indexing source; absolute Excel grid limits retained; warmup$Warmup;N$Iterations;invocations$Invocations;unroll1;all outliers retained; setup validates each value/type/schema";Cases=$caseSpecs;Reports=$reports;Mask=$(if($IsWindows){$Mask}else{'OS scheduled'})}|ConvertTo-Json -Depth 26|Set-Content -LiteralPath (Join-Path $root 'retained-native-dimensionless-v2.json') -Encoding utf8
