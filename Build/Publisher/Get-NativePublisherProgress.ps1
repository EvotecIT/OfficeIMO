function Get-NativePublisherProgress {
    <#
    .SYNOPSIS
    Reads the latest valid phase sequence belonging to one isolated worker.
    .PARAMETER ReceiptPath
    The process receipt and the prefix of its pre-identification startup receipt.
    .PARAMETER WorkerProcessId
    The exact controller-created worker PID. Foreign or incomplete data is ignored.
    .OUTPUTS
    A phase observation, or nothing when neither receipt can establish progress.
    #>
    param([string] $ReceiptPath, [int] $WorkerProcessId)

    foreach ($path in @($ReceiptPath, ($ReceiptPath + '.startup.json'))) {
        try {
            $receipt = Read-NativePublisherReceipt -Path $path
            [long] $sequence = 0
            if ($receipt.workerProcessId -eq $WorkerProcessId -and
                [long]::TryParse([string] $receipt.phaseSequence, [ref] $sequence) -and $sequence -gt 0 -and
                -not [string]::IsNullOrWhiteSpace($receipt.phase)) {
                return [pscustomobject]@{ sequence = $sequence; phase = [string] $receipt.phase }
            }
        } catch {
            # Partial startup writes and absent receipts do not establish progress.
        }
    }
}
