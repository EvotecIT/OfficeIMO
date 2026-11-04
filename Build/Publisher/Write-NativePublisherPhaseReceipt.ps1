function Write-NativePublisherPhaseReceipt {
    <#
    .SYNOPSIS
    Atomically updates an existing native-operation process receipt.
    .DESCRIPTION
    Writes the next JSON version to a temporary sibling, then replaces the receipt.
    Retries only Windows ERROR_SHARING_VIOLATION (HRESULT 0x80070020), for at most
    ten replacement attempts separated by 25 milliseconds. Other errors, or a
    sharing violation that remains after the final attempt, propagate to the
    worker's existing phase-failure handling. Native identity is not retried.
    .PARAMETER Path
    The existing process receipt created after native process identification.
    .PARAMETER Value
    The next phase receipt object.
    .OUTPUTS
    None. The old receipt remains intact when replacement fails.
    #>
    param(
        [Parameter(Mandatory)]
        [string] $Path,
        [Parameter(Mandatory)]
        $Value
    )

    $temporary = $Path + '.tmp'
    [IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try {
            [IO.File]::Replace($temporary, $Path, [Management.Automation.Language.NullString]::Value)
            return
        } catch {
            $failure = $_.Exception.GetBaseException()
            if ($failure -isnot [IO.IOException] -or $failure.HResult -ne -2147024864 -or $attempt -eq 10) {
                throw
            }
            [Threading.Thread]::Sleep(25)
        }
    }
}
