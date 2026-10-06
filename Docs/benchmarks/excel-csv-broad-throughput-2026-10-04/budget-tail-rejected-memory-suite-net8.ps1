param(
 [Parameter(Mandatory)][string]$CompareDirectory,
 [Parameter(Mandatory)][string]$PowerForgeAssembly,
 [Parameter(Mandatory)][string]$Worker,
 [string]$DotNet='dotnet',
 [long]$Mask=65535,
 [int]$Iterations=3
)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$manifest=Get-Content -LiteralPath (Join-Path $root 'manifest-net8.0.json') -Raw|ConvertFrom-Json
$toolHash=(Get-FileHash -LiteralPath $PowerForgeAssembly).Hash
$workerHash=(Get-FileHash -LiteralPath $Worker).Hash
function Assert-Frozen {
 foreach($file in $manifest.Manifest){foreach($side in 'Before','After'){
  $directory=if($side -eq 'Before'){'baseline-excel'}else{'candidate-excel'}
  if((Get-FileHash -LiteralPath (Join-Path $root "net8.0/$directory/$($file.Name)")).Hash -ne $file.$side){throw 'Memory snapshot changed'}
 }}
 if((Get-FileHash -LiteralPath $PowerForgeAssembly).Hash -ne $toolHash -or (Get-FileHash -LiteralPath $Worker).Hash -ne $workerHash){throw 'Memory tool changed'}
}
Assert-Frozen
if(!(Test-Path -LiteralPath (Join-Path $root 'memory-fixtures'))){throw 'Generate the independently validated common fixtures first'}
$outputRoot=Join-Path $root "memory-used-range-budget-v1-net8-$Mask-N$Iterations"
if(Test-Path -LiteralPath $outputRoot){throw 'Refuse existing memory measurements'}
New-Item -ItemType Directory -Path $outputRoot|Out-Null
$cases=@((Get-Content -LiteralPath (Join-Path $root 'memory-cases.json') -Raw|ConvertFrom-Json).PSObject.Properties.Name)
$validations=@();$samples=@()
foreach($case in $cases){foreach($engine in 'Before','After'){
 $output=Join-Path $outputRoot "$case-$engine-validate.json"
 & $DotNet $Worker $root $CompareDirectory $PowerForgeAssembly $case $engine $output $Mask validate *> (Join-Path $outputRoot "$case-$engine-validate.log")
 if($LASTEXITCODE){Get-Content -LiteralPath (Join-Path $outputRoot "$case-$engine-validate.log");throw 'Actual .NET8 complete-field validation failed'}
 $observation=Get-Content -LiteralPath $output -Raw|ConvertFrom-Json
 if(!$observation.CompleteFieldValidation -or $observation.Runtime -notmatch '^\.NET 8\.') {throw 'Wrong validation contract/runtime'}
 $validations+=$observation
}}
Write-Output "Actual .NET8 complete-field validations: $($validations.Count) separate workers"
for($iteration=0;$iteration -lt $Iterations;$iteration++){
 foreach($case in $cases){
  $engines=if($iteration%2 -eq 0){@('Before','After')}else{@('After','Before')}
  foreach($engine in $engines){
   $output=Join-Path $outputRoot "$case-$engine-$iteration.json"
   & $DotNet $Worker $root $CompareDirectory $PowerForgeAssembly $case $engine $output $Mask measure *> (Join-Path $outputRoot "$case-$engine-$iteration.log")
   if($LASTEXITCODE){Get-Content -LiteralPath (Join-Path $outputRoot "$case-$engine-$iteration.log");throw 'Actual .NET8 memory worker failed'}
   $sample=Get-Content -LiteralPath $output -Raw|ConvertFrom-Json
   if($sample.Runtime -notmatch '^\.NET 8\.' -or $sample.PowerForgeSha256 -ne $toolHash){throw 'Wrong memory runtime/tool'}
   $sample|Add-Member Iteration $iteration;$samples+=$sample
  }
 }
 Write-Output "Actual .NET8 memory iteration complete: $iteration"
}
if($samples.Count -ne 2*$cases.Count*$Iterations -or $validations.Count -ne 2*$cases.Count){throw 'Memory worker inventory incomplete'}
Assert-Frozen
[ordered]@{
 Manifest=$manifest;PowerForgeSha256=$toolHash;WorkerSha256=$workerHash
 Policy='One first complete read per fresh actual .NET8 worker; complete-field validation in separate workers; alternating Before/After order; no warmup or outlier removal; canonical PowerForge sampler'
 Validations=$validations;Samples=$samples
 RunnerSha256=(Get-FileHash -LiteralPath $PSCommandPath).Hash
}|ConvertTo-Json -Depth 22|Set-Content -LiteralPath (Join-Path $root "retained-memory-used-range-budget-v1-net8-$Mask-N$Iterations.json") -Encoding utf8
Write-Output "Actual .NET8 memory observations qualified: $($samples.Count) fresh workers"
