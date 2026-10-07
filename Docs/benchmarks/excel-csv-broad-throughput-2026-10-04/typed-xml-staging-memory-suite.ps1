param([Parameter(Mandatory)][string]$CompareDirectory,[Parameter(Mandatory)][string]$PowerForgeAssembly,[long]$Mask=65535,[int]$Iterations=3)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot;$pwsh=(Get-Process -Id $PID).Path
$manifest=Get-Content -LiteralPath (Join-Path $root 'manifest-net10.0.json') -Raw|ConvertFrom-Json
function Assert-Frozen {
 foreach($file in $manifest.Manifest){foreach($side in 'Before','After'){$directory=if($side -eq 'Before'){'baseline-excel'}else{'candidate-excel'};if((Get-FileHash -LiteralPath (Join-Path $root "net10.0/$directory/$($file.Name)")).Hash -ne $file.$side){throw 'Memory snapshot changed'}}}
}
Assert-Frozen
if(!(Test-Path -LiteralPath (Join-Path $root 'memory-fixtures'))){& $pwsh -NonInteractive -NoLogo -NoProfile -File (Join-Path $root 'memory-fixtures.ps1') -CompareDirectory $CompareDirectory;if($LASTEXITCODE){throw 'Memory fixture validation failed'}}
$outputRoot=Join-Path $root "memory-v1-$Mask-N$Iterations"
if(Test-Path -LiteralPath $outputRoot){throw 'Refuse existing memory measurements'}
New-Item -ItemType Directory -Path $outputRoot|Out-Null
$samples=@()
for($iteration=0;$iteration -lt $Iterations;$iteration++){
 foreach($case in @((Get-Content -LiteralPath (Join-Path $root 'cases.json') -Raw|ConvertFrom-Json).PSObject.Properties.Name)){
  $engines=if($iteration%2 -eq 0){@('Before','After')}else{@('After','Before')}
  foreach($engine in $engines){
   $output=Join-Path $outputRoot "$case-$engine-$iteration.json"
   & $pwsh -NonInteractive -NoLogo -NoProfile -File (Join-Path $root 'memory-worker.ps1') -CompareDirectory $CompareDirectory -PowerForgeAssembly $PowerForgeAssembly -Case $case -Engine $engine -Output $output -Mask $Mask *> (Join-Path $outputRoot "$case-$engine-$iteration.log")
   if($LASTEXITCODE){Get-Content -LiteralPath (Join-Path $outputRoot "$case-$engine-$iteration.log");throw 'Memory worker failed'}
   $sample=Get-Content -LiteralPath $output -Raw|ConvertFrom-Json;$sample|Add-Member Iteration $iteration;$samples+=$sample
  }
 }
 Write-Output "Typed memory iteration complete: $iteration"
}
if($samples.Count -ne 16*$Iterations){throw 'Memory worker inventory incomplete'}
Assert-Frozen
[ordered]@{Manifest=$manifest;Runner=@(foreach($name in 'memory-suite.ps1','memory-worker.ps1','memory-fixtures.ps1'){[pscustomobject]@{Name=$name;Sha256=(Get-FileHash -LiteralPath (Join-Path $root $name)).Hash}});Policy='One first complete typed read per fresh worker; alternate Before/After order; no warmup or outlier removal; after-return memory includes shared-pool retention and JIT/type initialization';Samples=$samples;Fixtures=@(Get-ChildItem -LiteralPath (Join-Path $root 'memory-fixtures') -Filter '*.json'|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json})}|ConvertTo-Json -Depth 22|Set-Content -LiteralPath (Join-Path $root "retained-memory-v1-$Mask-N$Iterations.json") -Encoding utf8
Write-Output "Typed memory observations qualified: $($samples.Count) fresh workers"