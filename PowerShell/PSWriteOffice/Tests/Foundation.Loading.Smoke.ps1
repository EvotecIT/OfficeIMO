param([Parameter(Mandatory)][string] $ModulePath, [Parameter(Mandatory)][string] $ConflictPath, [switch] $WithoutLoaderAssembly)
$ErrorActionPreference = 'Stop'
$priorCalls = [Collections.Generic.List[string]]::new()
$prior = [System.EventHandler[System.Management.Automation.CommandLookupEventArgs]]{ param($name,$eventArgs) $priorCalls.Add($name) }.GetNewClosure()
$ExecutionContext.InvokeCommand.CommandNotFoundAction = $prior
$conflict = [Reflection.Assembly]::LoadFrom($ConflictPath)
Import-Module $ModulePath -Force
if ($WithoutLoaderAssembly) {
    $module = Get-Module PSWriteOffice
    $registrationPath = Join-Path (Split-Path $ModulePath) 'Private/Register-OfficeCommandHints.ps1'
    & $module { param($path) $script:ModuleAssembly = $null; . $path } $registrationPath
}
$command = Get-Command New-OfficeWord -ErrorAction Stop
$context = [System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext($command.ImplementingType.Assembly)
$wordType = [OfficeIMO.Word.WordDocument]
$word = [OfficeIMO.Word.WordDocument]::Create()
try { $word.AddParagraph('Isolated document') | Out-Null } finally { $word.Dispose() }
$moduleOpenXml = @($context.Assemblies | Where-Object { $_.GetName().Name -eq 'DocumentFormat.OpenXml' })[0]
# Remove a legacy alias in this disposable process so command lookup exercises its actual replacement hint.
Remove-Item -LiteralPath Alias:ExcelNew -Force
$records = @(& { try { & 'ExcelNew' } catch { } } 3>&1)
$hint = @($records | Where-Object { $_ -is [Management.Automation.WarningRecord] } | ForEach-Object Message)
$priorWasCalled = $priorCalls.Contains('ExcelNew')
$handler = $ExecutionContext.InvokeCommand.CommandNotFoundAction
Remove-Module PSWriteOffice -Force
$restored = [object]::ReferenceEquals($ExecutionContext.InvokeCommand.CommandNotFoundAction, $prior)
[pscustomobject]@{
    WordType = $wordType.FullName
    WordIsolated = [object]::ReferenceEquals([System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext($wordType.Assembly), $context)
    ModuleIsDefault = [object]::ReferenceEquals($context, [System.Runtime.Loader.AssemblyLoadContext]::Default)
    ConflictVersion = $conflict.GetName().Version.ToString()
    ModuleOpenXmlVersion = $moduleOpenXml.GetName().Version.ToString()
    ConflictIsDefault = [object]::ReferenceEquals([System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext($conflict), [System.Runtime.Loader.AssemblyLoadContext]::Default)
    Hint = $hint
    PriorWasCalled = $priorWasCalled
    HandlerRestored = $restored
} | ConvertTo-Json -Depth 5 -Compress
