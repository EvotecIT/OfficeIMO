$ErrorActionPreference = 'Stop'

function Write-NotebookWorkerMessage {
    param($Writer, $Message)

    $Writer.WriteLine((ConvertTo-Json -InputObject $Message -Depth 3 -Compress))
}

function Limit-NotebookWorkerText {
    param([string] $Text, $State)

    if ($Text.Length -le $State.CharactersRemaining) {
        $State.CharactersRemaining -= $Text.Length
        return $Text
    }

    $length = $State.CharactersRemaining
    if ($length -gt 0 -and [char]::IsHighSurrogate($Text[$length - 1])) {
        $length--
    }
    $State.CharactersRemaining = 0
    $State.Truncated = $true
    return $Text.Substring(0, $length)
}

function Write-NotebookWorkerRecord {
    param($Writer, $Record, $State, [switch] $TerminatingError)

    $isErrorRecord = $Record -is [System.Management.Automation.ErrorRecord]
    $isNativeStandardError = $isErrorRecord -and -not $TerminatingError -and
        $Record.FullyQualifiedErrorId -in @('NativeCommandError', 'NativeCommandErrorMessage')
    $isError = $TerminatingError -or ($isErrorRecord -and -not $isNativeStandardError)
    if ($isError) {
        $State.Failed = $true
        if ($TerminatingError -and [object]::ReferenceEquals($State.LastError, $Record)) {
            return
        }
        $State.LastError = $Record
    }

    if ($State.RecordsRemaining -eq 0 -or $State.CharactersRemaining -eq 0) {
        $State.Truncated = $true
        return
    }

    $message = @{}
    if ($isError) {
        $message.kind = 'error'
        $text = if ($Record -is [Exception]) { $Record.Message } else { [string] $Record }
        $message.text = Limit-NotebookWorkerText $text $State
        if ($isErrorRecord) {
            $message.name = $Record.Exception.GetType().Name
            $traceback = [System.Collections.Generic.List[string]]::new()
            foreach ($frame in ($Record.ScriptStackTrace -split '\r?\n', 17)) {
                if ([string]::IsNullOrEmpty($frame)) {
                    continue
                }
                if ($traceback.Count -eq 16 -or $State.CharactersRemaining -eq 0) {
                    $State.Truncated = $true
                    break
                }
                $traceback.Add((Limit-NotebookWorkerText $frame $State))
            }
            $message.traceback = $traceback.ToArray()
        } else {
            $message.name = $Record.GetType().Name
        }
    } elseif ($Record -is [System.Management.Automation.InformationRecord]) {
        $message.kind = 'stdout'
        $text = [string] $Record.MessageData
        if ($Record.MessageData -isnot [System.Management.Automation.HostInformationMessage] -or
            -not $Record.MessageData.NoNewLine) {
            $text += "`n"
        }
        $message.text = Limit-NotebookWorkerText $text $State
    } elseif ($isNativeStandardError -or $Record -is [System.Management.Automation.InformationalRecord]) {
        $message.kind = 'stderr'
        $text = if ($isNativeStandardError) { [string] $Record } else { $Record.Message }
        $message.text = Limit-NotebookWorkerText ($text + "`n") $State
    } else {
        $message.kind = 'output'
        try {
            $text = [string] $Record
        } catch {
            $text = '[PowerShell could not render this result as text.]'
            $State.Truncated = $true
        }
        $message.text = Limit-NotebookWorkerText $text $State
    }

    $State.RecordsRemaining--
    Write-NotebookWorkerMessage $Writer $message
}

function Write-NotebookWorkerOutput {
    param($Writer, $OutputBuffer, $State)

    # Drain even after truncation so a noisy pipeline can finish and keep its session state.
    $records = $OutputBuffer.ReadAll()
    if ($records.Count -eq 0) {
        return
    }

    foreach ($record in $records) {
        if ($State.RecordsRemaining -eq 0 -or $State.CharactersRemaining -eq 0) {
            $State.Truncated = $true
            if ($State.Failed) {
                break
            }
            # Merged PowerShell errors can leave HadErrors unset even when the cell failed.
            if ($record -isnot [System.Management.Automation.ErrorRecord]) {
                continue
            }
        }
        Write-NotebookWorkerRecord $Writer $record $State
    }
    if ($State.Truncated -and -not $State.TruncationReported) {
        $State.TruncationReported = $true
        Write-NotebookWorkerMessage $Writer @{ kind = 'truncated' }
    }
    $Writer.Flush()
}

function Write-NotebookWorkerCompletion {
    param($Writer, [bool] $Failed, [bool] $Truncated)

    Write-NotebookWorkerMessage $Writer @{
        kind      = 'complete'
        failed    = $Failed
        truncated = $Truncated
    }
    $Writer.Flush()
}

function Invoke-NotebookWorkerCell {
    param($Writer, $Runspace, $BeginInvoke, $Request)

    if ([string]::IsNullOrWhiteSpace($Request.source)) {
        Write-NotebookWorkerCompletion $Writer $false $false
        return
    }

    $state = @{
        CharactersRemaining = [int] $Request.maxOutputCharacters
        RecordsRemaining    = [int] $Request.maxOutputRecords
        Failed              = $false
        Truncated           = $false
        TruncationReported  = $false
        LastError           = $null
    }
    $pipeline = [PowerShell]::Create()
    $output = [System.Management.Automation.PSDataCollection[psobject]]::new()
    try {
        $pipeline.Runspace = $Runspace
        [void] $pipeline.AddScript($Request.source, $false)
        $pipeline.Commands.Commands[0].MergeMyResults(
            [System.Management.Automation.Runspaces.PipelineResultTypes]::All,
            [System.Management.Automation.Runspaces.PipelineResultTypes]::Output)

        $invocation = $BeginInvoke.Invoke($pipeline, [object[]] @($null, $output))
        while (-not $invocation.IsCompleted) {
            Write-NotebookWorkerOutput $Writer $output $state
            if ($pipeline.Streams.Progress.Count -gt 0) {
                $pipeline.Streams.Progress.Clear()
            }
            [void] $invocation.AsyncWaitHandle.WaitOne(5)
        }

        Write-NotebookWorkerOutput $Writer $output $state
        try {
            [void] $pipeline.EndInvoke($invocation)
        } catch {
            $failure = $pipeline.InvocationStateInfo.Reason
            if ($null -eq $failure) {
                throw
            }
            $record = $failure
            if ($failure -is [System.Management.Automation.RuntimeException] -and $null -ne $failure.ErrorRecord) {
                $record = $failure.ErrorRecord
            }
            Write-NotebookWorkerRecord $Writer $record $state -TerminatingError
        }

        Write-NotebookWorkerCompletion $Writer ($state.Failed -or $pipeline.HadErrors) $state.Truncated
    } finally {
        $pipeline.Dispose()
        $output.Dispose()
    }
}

$pipe = $null
$reader = $null
$writer = $null
$runspace = $null
try {
    $pipe = [System.IO.Pipes.NamedPipeClientStream]::new(
        '.', $env:OFFICEIMO_NOTEBOOK_PIPE, [System.IO.Pipes.PipeDirection]::InOut)
    $pipe.Connect(30000)
    $env:OFFICEIMO_NOTEBOOK_PIPE = $null
    $encoding = [System.Text.UTF8Encoding]::new($false, $true)
    $reader = [System.IO.StreamReader]::new($pipe, $encoding, $false, 4096, $true)
    $writer = [System.IO.StreamWriter]::new($pipe, $encoding, 4096, $true)
    $writer.NewLine = "`n"

    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable('ProgressPreference', 'SilentlyContinue')

    # Windows PowerShell cannot call a generic method with explicit type arguments.
    $beginInvoke = [PowerShell].GetMethods() | Where-Object {
        $_.Name -eq 'BeginInvoke' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Length -eq 2
    }
    $beginInvoke = $beginInvoke.MakeGenericMethod([psobject], [psobject])

    Write-NotebookWorkerMessage $writer @{
        kind = 'ready'
        text = $PSVersionTable.PSVersion.ToString()
    }
    $writer.Flush()
    while ($null -ne ($line = $reader.ReadLine())) {
        $request = ConvertFrom-Json -InputObject $line
        if ($request.source -isnot [string] -or $request.maxOutputCharacters -lt 1 -or $request.maxOutputRecords -lt 1) {
            throw 'The notebook worker received an invalid request.'
        }
        Invoke-NotebookWorkerCell $writer $runspace $beginInvoke $request
    }
} catch {
    if ($null -ne $writer -and $pipe.IsConnected) {
        try {
            $failure = $_.Exception.Message
            Write-NotebookWorkerMessage $writer @{
                kind = 'fatal'
                text = $failure.Substring(0, [Math]::Min($failure.Length, 2048))
            }
            $writer.Flush()
        } catch [System.IO.IOException] {
            # The caller may have closed the transport while reporting the failure.
        }
    } else {
        [Console]::Error.WriteLine($_.Exception.Message)
    }
    exit 1
} finally {
    if ($null -ne $runspace) {
        $runspace.Dispose()
    }
    if ($null -ne $writer) {
        $writer.Dispose()
    }
    if ($null -ne $reader) {
        $reader.Dispose()
    }
    if ($null -ne $pipe) {
        $pipe.Dispose()
    }
}
