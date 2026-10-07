param([Parameter(Mandatory)][string]$CompareDirectory,[Parameter(Mandatory)][string]$PowerForgeAssembly,[Parameter(Mandatory)][string]$Case,[ValidateSet('Before','After')][string]$Engine,[Parameter(Mandatory)][string]$Output,[long]$Mask=65535)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$fixture=Get-Content -LiteralPath (Join-Path $root "memory-fixtures/$Case.json") -Raw|ConvertFrom-Json
$spec=(Get-Content -LiteralPath (Join-Path $root 'cases.json') -Raw|ConvertFrom-Json).$Case
[string]$path=Join-Path $root "memory-fixtures/$Case.xlsx"
if((Get-FileHash -LiteralPath $path).Hash -ne $fixture.FileSha256){throw 'Memory input changed'}
if($IsWindows -or $IsLinux){[Diagnostics.Process]::GetCurrentProcess().ProcessorAffinity=[IntPtr]::new($Mask)}
[Diagnostics.Process]::GetCurrentProcess().PriorityClass='Normal'
$harness=[Reflection.Assembly]::LoadFrom((Join-Path $CompareDirectory 'bin/Release/net10.0/Compare.dll'))
$contextType=$harness.GetType('SnapshotContext',$true)
$directory=Join-Path $root "net10.0/$(if($Engine -eq 'Before'){'baseline'}else{'candidate'})-excel"
$context=$contextType.GetConstructor([type[]]@([string])).Invoke([object[]]@([string]$directory))
$assembly=$context.LoadFromAssemblyPath((Join-Path $directory 'OfficeIMO.Excel.Benchmarks.dll'))
$type=$assembly.GetType('OfficeIMO.Excel.Benchmarks.'+$spec.Class,$true)
$instance=[Activator]::CreateInstance($type)
foreach($setting in $spec.Properties.PSObject.Properties){$property=$type.GetProperty($setting.Name);$value=[Convert]::ChangeType($setting.Value,$property.PropertyType,[Globalization.CultureInfo]::InvariantCulture);$property.SetValue($instance,$value)}
$flags=[Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic
$type.GetField('_path',$flags).SetValue($instance,$path)
foreach($setting in $fixture.State.PSObject.Properties){$field=$type.GetField($setting.Name,$flags);$field.SetValue($instance,[Convert]::ChangeType($setting.Value,$field.FieldType,[Globalization.CultureInfo]::InvariantCulture))}
$method=$type.GetMethod($spec.Method)
$probeStart=[Reflection.Assembly]::LoadFile($PowerForgeAssembly).GetType('PowerForge.BenchmarkMemoryProbe',$true).GetMethod('Start')
[GC]::Collect();[GC]::WaitForPendingFinalizers();$baseline=[GC]::GetTotalMemory($true)
$sample=$probeStart.Invoke($null,@([int]5))
try{
 $allocated=[GC]::GetAllocatedBytesForCurrentThread();$allThreadAllocated=[GC]::GetTotalAllocatedBytes($true);$timer=[Diagnostics.Stopwatch]::StartNew()
 $observed=$method.Invoke($instance,$null)
 $timer.Stop();$totalAllocation=[GC]::GetTotalAllocatedBytes($true)-$allThreadAllocated;$allocation=[GC]::GetAllocatedBytesForCurrentThread()-$allocated
 if($observed.ToString() -ne $fixture.Observation){throw 'Complete mapped result differs from validated producer'}
 $peak=$sample.Complete();$retained=[GC]::GetTotalMemory($true)-$baseline
 [ordered]@{Case=$Case;Engine=$Engine;Pid=$PID;Runtime=[Runtime.InteropServices.RuntimeInformation]::FrameworkDescription;Host=[Environment]::MachineName;Mask=$(if($IsWindows){$Mask}else{'OS scheduled'});FileSha256=$fixture.FileSha256;ExcelDllSha256=(Get-FileHash -LiteralPath (Join-Path $directory 'OfficeIMO.Excel.dll')).Hash;ExpectedObservation=$fixture.Observation;Observed=$observed;ReadMs=$timer.Elapsed.TotalMilliseconds;CallingThreadAllocatedBytes=$allocation;AllThreadAllocatedBytesIncludingSampler=$totalAllocation;AfterReturnManagedIncreaseBytes=$retained;Peak=$peak;Contract='First complete typed or tabular read in a fresh worker; producer separately validated every projected field; readers and materialized result released by benchmark before retained measurement';Limits='JIT/type initialization and reflection invocation included; calling-thread allocation excludes sampler and prefetch worker; all-thread allocation includes sampler; sampled peaks are lower bounds; operating system caches warmed by producer; no warmed throughput claim'}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $Output -Encoding utf8
}finally{$sample.Dispose()}
