$script:Workers = @{}

function Invoke-NotebookWorker {
    param(
        [string] $Session,
        [string] $Name,
        [string] $Executable,
        [string[]] $Arguments,
        [hashtable] $Request,
        [int] $TimeoutSeconds
    )
    $key = "$Session/$Name"
    if (-not $script:Workers.ContainsKey($key)) {
        $start = [Diagnostics.ProcessStartInfo]::new($Executable)
        $start.UseShellExecute = $false
        $start.RedirectStandardInput = $true
        $start.RedirectStandardOutput = $true
        $start.StandardInputEncoding = [Text.UTF8Encoding]::new($false)
        $start.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
        foreach ($argument in $Arguments) { $start.ArgumentList.Add($argument) }
        $script:Workers[$key] = [Diagnostics.Process]::Start($start)
    }
    $process = $script:Workers[$key]
    $process.StandardInput.WriteLine(($Request | ConvertTo-Json -Compress))
    $process.StandardInput.Flush()
    $line = $process.StandardOutput.ReadLineAsync().WaitAsync([TimeSpan]::FromSeconds($TimeoutSeconds)).GetAwaiter().GetResult()
    if ($null -eq $line) { throw "$Name benchmark worker exited before replying." }
    $response = $line | ConvertFrom-Json
    if (-not $response.ok) { throw $response.error }
    $response
}

function Stop-NotebookWorkers {
    param([string] $Session, [string] $Name)
    $keys = @($script:Workers.Keys | Where-Object { $_.StartsWith("$Session/") -and (-not $Name -or $_ -eq "$Session/$Name") })
    foreach ($key in $keys) {
        $process = $script:Workers[$key]
        try {
            if (-not $process.HasExited) {
                $process.StandardInput.Close()
                if (-not $process.WaitForExit(3000)) {
                    $process.Kill($true)
                    [void] $process.WaitForExit(3000)
                }
            }
        } finally {
            $script:Workers.Remove($key)
            $process.Dispose()
        }
    }
}

Export-ModuleMember -Function Invoke-NotebookWorker, Stop-NotebookWorkers
