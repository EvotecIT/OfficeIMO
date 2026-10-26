# Runs in the module scope after its binary has loaded. Preserve unrelated lookup hooks.
$officeHintStream = $ModuleAssembly.GetManifestResourceStream('PSWriteOffice.CommandHints.json')
if ($null -eq $officeHintStream) { throw 'The PSWriteOffice command-hint resource is missing. Rebuild the module.' }
$officeHintReader = [IO.StreamReader]::new($officeHintStream)
try { $officeHintData = $officeHintReader.ReadToEnd() | ConvertFrom-Json } finally { $officeHintReader.Dispose() }
$officeHints = @{}
foreach ($entry in $officeHintData.PSObject.Properties) { $officeHints[$entry.Name] = $entry.Value }
$officeInvokeCommand = $ExecutionContext.InvokeCommand
$officePreviousLookup = $officeInvokeCommand.CommandNotFoundAction
$officeRegistration = @{ Active = $true }
$officeLookup = [System.EventHandler[System.Management.Automation.CommandLookupEventArgs]]{
    param($name, $eventArgs)
    if ($null -ne $officePreviousLookup) { $officePreviousLookup.Invoke($name, $eventArgs) }
    if ($officeRegistration.Active -and $null -eq $eventArgs.Command -and $officeHints.ContainsKey($name)) {
        Microsoft.PowerShell.Utility\Write-Warning "$name was replaced by $($officeHints[$name])"
    }
}.GetNewClosure()
$officeInvokeCommand.CommandNotFoundAction = $officeLookup
$officePreviousRemoval = $ExecutionContext.SessionState.Module.OnRemove
$ExecutionContext.SessionState.Module.OnRemove = {
    try {
        $officeRegistration.Active = $false
        if ([object]::ReferenceEquals($officeInvokeCommand.CommandNotFoundAction, $officeLookup)) {
            $officeInvokeCommand.CommandNotFoundAction = $officePreviousLookup
        }
    } finally {
        if ($null -ne $officePreviousRemoval) { & $officePreviousRemoval @args }
    }
}.GetNewClosure()
