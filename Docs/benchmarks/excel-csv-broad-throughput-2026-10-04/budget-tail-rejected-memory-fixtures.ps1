param([Parameter(Mandatory)][string]$CompareDirectory)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$env:OFFICEIMO_PERFORMANCE_SNAPSHOTS=Join-Path $root 'net10.0'
$env:OFFICEIMO_COMPARISON_CASES=Join-Path $root 'memory-cases.json'
$env:OFFICEIMO_BENCHMARK_DATA=Join-Path $root 'memory-producer'
$env:OFFICEIMO_PERFORMANCE_AA='0'
[Reflection.Assembly]::LoadFrom((Join-Path $CompareDirectory 'bin/Release/net10.0/Compare.dll'))|Out-Null
$destination=Join-Path $root 'memory-fixtures'
if(Test-Path -LiteralPath $destination){throw 'Refuse existing memory fixtures'}
New-Item -ItemType Directory -Path $destination|Out-Null
$flags=[Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
foreach($key in [SnapshotProbe]::Keys()){
 $before=[SnapshotProbe]::Load($key,'baseline');$after=$null
 try{
  $after=[SnapshotProbe]::Load($key,'candidate');[SnapshotProbe]::ValidateEquivalent($before,$after)
  $inputs=@{};$partLengths=@{};$state=@{}
  foreach($side in 'Before','After'){
   $probe=if($side -eq 'Before'){$before}else{$after}
   $instance=[SnapshotProbe].GetField('_instance',$flags).GetValue($probe)
   $type=$instance.GetType();$input=[string]$type.GetField('_path',$flags).GetValue($instance)
   $parts=[ordered]@{};$zip=[IO.Compression.ZipFile]::OpenRead($input)
   try{foreach($part in $zip.Entries|Where-Object {$_.FullName -match '^xl/(worksheets/|styles.xml$|sharedStrings.xml$)'}|Sort-Object FullName){$partLengths[$part.FullName]=$part.Length;$stream=$part.Open();$hash=[Security.Cryptography.SHA256]::Create();try{$parts[$part.FullName]=[Convert]::ToHexString($hash.ComputeHash($stream))}finally{$hash.Dispose();$stream.Dispose()}}}finally{$zip.Dispose()}
   $inputs[$side]=$parts
   if($side -eq 'Before'){
    Copy-Item -LiteralPath $input -Destination (Join-Path $destination "$key.xlsx")
    foreach($field in $type.GetFields($flags)){
     if($field.FieldType -in [string],[int],[long] -and $field.Name -ne '_path'){$state[$field.Name]=$field.GetValue($instance)}
    }
    $observation=$probe.RunSynchronously().ToString()
   }
  }
  if($inputs.Before.Count -ne $inputs.After.Count){throw 'Memory fixture part inventory differs'}
  foreach($part in $inputs.Before.Keys){if($inputs.Before[$part] -ne $inputs.After[$part]){throw 'Memory fixture part differs'}}
  [ordered]@{Case=$key;FileSha256=(Get-FileHash -LiteralPath (Join-Path $destination "$key.xlsx")).Hash;Parts=$inputs;PartLengths=$partLengths;State=$state;Observation=$observation;Validation='Both setup routines validate every projected field of every row; complete observations equal; decoded part hashes identical';ProducerProcess=$PID}|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $destination "$key.json") -Encoding utf8
 }finally{$before.Dispose();if($after){$after.Dispose()}}
}
Write-Output 'Memory fixtures independently validated'
