function Read-NativePublisherReceipt {
    <#
    .SYNOPSIS
    Reads one native-operation receipt without blocking atomic replacement.
    .DESCRIPTION
    Permits the worker to replace its receipt while this reader holds the previous
    file version. Releases the file handle before parsing JSON. Missing, partial or
    malformed receipts throw; callers decide whether the operation can still wait.
    .PARAMETER Path
    The process receipt written by the isolated native worker.
    .OUTPUTS
    The decoded receipt object.
    #>
    param([Parameter(Mandatory)][string] $Path)

    $sharing = [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, $sharing)
    $reader = $null
    try {
        $reader = [IO.StreamReader]::new($stream)
        $json = $reader.ReadToEnd()
    } finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
        $stream.Dispose()
    }
    $json | ConvertFrom-Json -ErrorAction Stop
}
