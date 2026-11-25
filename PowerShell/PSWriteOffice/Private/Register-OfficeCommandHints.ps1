# Runs in the module scope after its binary has loaded. Preserve unrelated lookup hooks.
# Import-Module registers the binary module's own types on both editions. Avoid Core-only loader variables.
$officeHintAssembly = [PSWriteOffice.Initialize].Assembly
$officeHintStream = $officeHintAssembly.GetManifestResourceStream('PSWriteOffice.CommandHints.json')
if ($null -eq $officeHintStream) { throw 'The PSWriteOffice command-hint resource is missing. Rebuild the module.' }
$officeHintReader = [IO.StreamReader]::new($officeHintStream)
try { $officeHintText = $officeHintReader.ReadToEnd() } finally { $officeHintReader.Dispose() }
# Parse the hints on the first lookup that names one. Import then needs no JSON cmdlets.
$officeRegistration = @{ Active = $true; Text = $officeHintText; Hints = $null }
$officeInvokeCommand = $ExecutionContext.InvokeCommand
$officePreviousLookup = $officeInvokeCommand.CommandNotFoundAction
$officeLookup = [System.EventHandler[System.Management.Automation.CommandLookupEventArgs]]{
    param($name, $eventArgs)
    if ($null -ne $officePreviousLookup) { $officePreviousLookup.Invoke($name, $eventArgs) }
    if (-not $officeRegistration.Active -or $null -ne $eventArgs.Command) { return }
    if ($null -eq $officeRegistration.Hints) {
        if ($officeRegistration.Text.IndexOf('"' + $name + '"', [StringComparison]::OrdinalIgnoreCase) -lt 0) { return }
        $hints = @{}
        foreach ($entry in (Microsoft.PowerShell.Utility\ConvertFrom-Json -InputObject $officeRegistration.Text).PSObject.Properties) { $hints[$entry.Name] = $entry.Value }
        $officeRegistration.Hints = $hints
    }
    if ($officeRegistration.Hints.ContainsKey($name)) {
        Microsoft.PowerShell.Utility\Write-Warning "$name is no longer available. $($officeRegistration.Hints[$name])"
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
