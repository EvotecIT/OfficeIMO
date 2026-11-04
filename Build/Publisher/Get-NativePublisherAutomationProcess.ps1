function Get-NativePublisherAutomationProcess {
    <#
    .SYNOPSIS
    Identifies the sole newly activated Publisher automation server before a document is opened.
    .DESCRIPTION
    The worker must first establish an empty Publisher process baseline. A new,
    same-session process must expose both automation-server command-line flags,
    its executable path and a matching process start time. Ambiguity refuses
    ownership; no process is stopped here.
    .PARAMETER StartedAfterUtc
    The timestamp captured before this worker activated the COM application.
    .OUTPUTS
    The verified process handle, retained and disposed by the isolated worker.
    #>
    param([Parameter(Mandatory)][datetime] $StartedAfterUtc)

    $servers = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='MSPUB.EXE'" -ErrorAction Stop)
    if ($servers.Count -ne 1) {
        throw 'COM activation did not identify one isolated Publisher automation server.'
    }
    $server = $servers[0]
    if ([string]::IsNullOrWhiteSpace($server.ExecutablePath) -or
        $server.CommandLine -notmatch '(?i)(?:^|\s)/Automation(?:\s|$)' -or
        $server.CommandLine -notmatch '(?i)(?:^|\s)-Embedding(?:\s|$)' -or
        $server.CreationDate.ToUniversalTime() -lt $StartedAfterUtc.ToUniversalTime()) {
        throw 'The new Publisher process has no qualified automation-server identity.'
    }
    $process = Get-Process -Id $server.ProcessId -ErrorAction Stop
    if ($process.ProcessName -ne 'MSPUB' -or $process.Path -ine $server.ExecutablePath -or
        $process.SessionId -ne (Get-Process -Id $PID).SessionId -or
        [Math]::Abs(($process.StartTime.ToUniversalTime() - $server.CreationDate.ToUniversalTime()).TotalMilliseconds) -gt 1) {
        $process.Dispose()
        throw 'The Publisher process changed or belongs to another session; ownership is unqualified.'
    }
    return $process
}
