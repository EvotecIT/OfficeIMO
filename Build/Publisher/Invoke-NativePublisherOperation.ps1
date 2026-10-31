<#
.SYNOPSIS
Runs one bounded, isolated native Publisher qualification operation.
.DESCRIPTION
Shared process owner for comparison-corpus production and candidate validation.
The worker disables macros and identifies its own Publisher process through the
COM window before authoring or opening the candidate. Existing Publisher sessions
are preserved. A failed startup without an identity receipt never authorizes
stopping a process by its name or start time alone.
.PARAMETER OutputDirectory
Task-owned directory for generated evidence. Existing evidence is never overwritten.
.PARAMETER Operation
Corpus produces one synthetic native case. Validate opens a byte-identical copy
of InputPath read-only and performs the selected Stage. Full remains the default:
it inventories and exports the input, then saves, reopens and exports a separate PUB.
.PARAMETER Case
Synthetic case selected by New-NativeComparisonCorpus.ps1.
.PARAMETER InputPath
OfficeIMO-generated candidate to validate. Required for Validate.
.PARAMETER Stage
Inspect performs read-only open and inventory. Resave also saves and reopens a
separate PUB without PDF export. Full performs both PDF exports as well. Partial
stages establish only their listed operations; native repair remains unassessed.
.PARAMETER TimeoutSeconds
Bounded worker lifetime, including application startup and shutdown.
.PARAMETER NativeEditText
Optional single-line replacement for a candidate with one physical page and one
unlinked text box. Validate/Full only. The writable snapshot is changed before
native save-as; initial artifacts retain the original content. An empty value
clears the text range.
.EXAMPLE
./Build/Publisher/Invoke-NativePublisherOperation.ps1 -Operation Validate -InputPath ./candidate.pub -OutputDirectory ./candidate-native
.EXAMPLE
./Build/Publisher/Invoke-NativePublisherOperation.ps1 -Operation Validate -Stage Inspect -InputPath ./candidate.pub -OutputDirectory ./candidate-inspect
.OUTPUTS
The native case or validation report. Native repair status remains unassessed
unless an independent native UI observation supplies that evidence.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputDirectory,

    [ValidateSet('Corpus', 'Validate')]
    [string] $Operation = 'Corpus',

    [ValidateSet('Blank', 'BlankPages', 'Text', 'TextLong', 'TextBlank', 'TextParagraphs', 'TextNonuniformParagraphs', 'TextNonuniformTextReplacements', 'TextNonuniformEmptyParagraphs', 'TextUniformEmptyParagraphs', 'TextParagraphAttributes', 'TextAlignedTabs', 'TextFrameLayout', 'TextStories', 'TextFormattingPages', 'TextFormattingMany', 'TextCharacterPages', 'TextCharacterAttributes', 'TextUnderlinePatterns', 'ZeroAnchors', 'Rectangle', 'Image', 'Newsletter', 'TableLabel', 'TableLabelWrapped', 'TableLabelShortParagraphs', 'TableCellTyping', 'TableCellTypingEarlyRange', 'TableCellParagraphs', 'TableNonuniformParagraphs', 'TableNonuniformTextReplacements', 'TableNonuniformEmptyParagraphs', 'TableUniformEmptyParagraphs', 'Brochure', 'MultipleMasters', 'TwoPageMaster', 'PictureOwnership', 'PictureReplacement', 'PictureCropped')]
    [string] $Case = 'Blank',

    [string] $InputPath,

    [ValidateSet('Inspect', 'Resave', 'Full')]
    [string] $Stage = 'Full',

    [AllowEmptyString()]
    [ValidateLength(0, 4096)]
    [ValidatePattern('\A[^\p{Cc}\p{Zl}\p{Zp}]*\z')]
    [string] $NativeEditText,

    [ValidateRange(10, 600)]
    [int] $TimeoutSeconds = 90
)

$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') {
    throw 'Native qualification requires Microsoft Publisher on Windows.'
}
$nativeEditRequested = $PSBoundParameters.ContainsKey('NativeEditText')
if ($nativeEditRequested -and ($Operation -ne 'Validate' -or $Stage -ne 'Full')) {
    throw 'NativeEditText requires candidate validation with Stage Full.'
}
if ($Operation -eq 'Validate') {
    if ([string]::IsNullOrWhiteSpace($InputPath) -or -not (Test-Path -LiteralPath $InputPath -PathType Leaf)) {
        throw 'Validate requires an existing native PUB candidate.'
    }
    $InputPath = [IO.Path]::GetFullPath($InputPath)
    if ([IO.Path]::GetExtension($InputPath) -ne '.pub') {
        throw 'Validate accepts a native .pub candidate only.'
    }
} elseif ($Stage -ne 'Full') {
    throw 'Inspect and Resave stages apply only to candidate validation.'
}
. (Join-Path $PSScriptRoot 'Get-NativeValidationArtifactNames.ps1')
. (Join-Path $PSScriptRoot 'Read-NativePublisherReceipt.ps1')
$output = [IO.Path]::GetFullPath($OutputDirectory)
$runDirectory = Join-Path $output '.run'
$key = $Operation + '-' + $Case
if ($Operation -eq 'Validate') {
    $key = 'Validate-Candidate'
}
$receiptPath = Join-Path $runDirectory ($key + '.process.json')
$stdout = Join-Path $runDirectory ($key + '.stdout.txt')
$stderr = Join-Path $runDirectory ($key + '.stderr.txt')
$reservationPath = Join-Path $runDirectory ($key + '.reservation.json')
$diagnosticPaths = @(
    $receiptPath,
    $stdout,
    $stderr,
    $reservationPath,
    (Join-Path $runDirectory ($key + '.closure.json')),
    (Join-Path $runDirectory ($key + '.failure.json')),
    (Join-Path $runDirectory ($key + '.cleanup.json')),
    ($receiptPath + '.startup.json'),
    ($receiptPath + '.tmp')
)
$artifactNames = @(Get-NativeValidationArtifactNames -Stage $Stage)
if ($Operation -eq 'Corpus') {
    $stem = 'native-' + $Case.ToLowerInvariant()
    $artifactNames = @(($stem + '.pub'), ($stem + '.pdf'), ($stem + '.json'))
    if ($Case -in @('Image', 'Newsletter', 'TableLabel', 'TableLabelWrapped', 'TableLabelShortParagraphs', 'PictureOwnership', 'PictureReplacement', 'PictureCropped')) {
        $artifactNames += $stem + '.png'
        if ($Case -eq 'PictureReplacement') {
            $artifactNames += $stem + '-replacement.png'
        }
    }
}
$plannedPaths = @($artifactNames | ForEach-Object { Join-Path $output $_ }) + $diagnosticPaths
foreach ($path in $plannedPaths) {
    if (Test-Path -LiteralPath $path) {
        throw ('Existing operation evidence is not overwritten: ' + $path)
    }
}
if (@(Get-Process -Name MSPUB -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'Publisher is already running. Preserve that session and run this tool after it is closed.'
}
[void][IO.Directory]::CreateDirectory($runDirectory)
$reservation = [ordered]@{
    schema = 'officeimo.publisher.native-reservation.v1'
    controllerProcessId = $PID
    operation = $Operation
    case = $Case
    artifacts = $artifactNames
}
if ($Operation -eq 'Validate') {
    $reservation.stage = $Stage
}
$reservationBytes = [Text.UTF8Encoding]::new($false).GetBytes(($reservation | ConvertTo-Json -Depth 4))
$reservationStream = [IO.FileStream]::new($reservationPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
try {
    $reservationStream.Write($reservationBytes, 0, $reservationBytes.Length)
} finally {
    $reservationStream.Dispose()
}

function ConvertTo-NativePublisherArgument {
    <# Quote one Windows argv value, including embedded quotes and trailing backslashes. #>
    param([AllowEmptyString()][string] $Argument)
    $quoted = [Text.StringBuilder]::new()
    [void]$quoted.Append([char]34)
    $backslashes = 0
    foreach ($character in $Argument.ToCharArray()) {
        if ($character -eq [char]92) {
            $backslashes++
            continue
        }
        if ($character -eq [char]34) {
            [void]$quoted.Append([char]92, ($backslashes * 2 + 1))
        } else {
            [void]$quoted.Append([char]92, $backslashes)
        }
        [void]$quoted.Append($character)
        $backslashes = 0
    }
    [void]$quoted.Append([char]92, ($backslashes * 2))
    [void]$quoted.Append([char]34)
    return $quoted.ToString()
}

$workerPath = Join-Path $PSScriptRoot 'Invoke-NativeComparisonCase.ps1'
$hostPath = (Get-Process -Id $PID).Path
$controllerStart = (Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks
$argumentValues = @(
    '-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-File', $workerPath,
    '-OutputDirectory', $output, '-Operation', $Operation, '-Case', $Case,
    '-ReceiptPath', $receiptPath, '-ControllerProcessId', $PID.ToString([Globalization.CultureInfo]::InvariantCulture),
    '-ControllerStartTimeUtcTicks', $controllerStart.ToString([Globalization.CultureInfo]::InvariantCulture)
)
if ($Operation -eq 'Validate') {
    $argumentValues += @('-InputPath', $InputPath, '-Stage', $Stage)
}
if ($nativeEditRequested) {
    $argumentValues += @('-NativeEditText', $NativeEditText)
}
$arguments = ($argumentValues | ForEach-Object { ConvertTo-NativePublisherArgument $_ }) -join ' '
$worker = $null
$succeeded = $false
$report = $null
$primaryError = $null
$cleanupError = $null
$closure = [ordered]@{
    schema = 'officeimo.publisher.native-process-closure.v1'
    operation = $Operation
    case = $Case
    workerStopped = $false
    workerTeardown = 'Exited'
    native = 'NotIdentified'
    error = $null
}

function Get-CompletionReceipt {
    param([int] $WorkerProcessId)
    if (-not (Test-Path -LiteralPath $receiptPath)) {
        return $null
    }
    try {
        $receipt = Read-NativePublisherReceipt -Path $receiptPath
        if ($receipt.workerProcessId -eq $WorkerProcessId -and $receipt.phase -eq 'ShutdownComplete' -and $receipt.operationSucceeded -is [bool]) {
            return $receipt
        }
    } catch {
        # A startup receipt can be between its first write and replacement.
        # Malformed/incomplete data never qualifies as operation completion.
    }
    return $null
}

try {
    $worker = Start-Process -FilePath $hostPath -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    # Windows PowerShell needs the owned process handle retained before timed
    # waits; reacquiring a short-lived PID after exit can lose its exit code.
    [void]$worker.Handle
    $workerStartTicks = $worker.StartTime.ToUniversalTime().Ticks
    $elapsed = [Diagnostics.Stopwatch]::StartNew()
    $completion = $null
    $workerExited = $false
    while ($elapsed.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if ($worker.WaitForExit(200)) {
            $workerExited = $true
            $completion = Get-CompletionReceipt -WorkerProcessId $worker.Id
            break
        }
        $completion = Get-CompletionReceipt -WorkerProcessId $worker.Id
        if ($null -ne $completion) {
            # Explicit completion is emitted only after evidence writes and
            # Close/Quit. COM host teardown can still block even Environment.Exit.
            if (-not $worker.WaitForExit(1000)) {
                if ($worker.StartTime.ToUniversalTime().Ticks -ne $workerStartTicks) {
                    throw 'The worker process identity changed before completion cleanup.'
                }
                $worker.Kill()
                if (-not $worker.WaitForExit(10000)) {
                    throw 'The completed native-operation worker did not exit after bounded cleanup.'
                }
                $closure.workerStopped = $true
                $closure.workerTeardown = 'StoppedAfterCompletionReceipt'
            }
            $workerExited = $true
            break
        }
    }
    if (-not $workerExited) {
        $closure.workerTeardown = 'StoppedAfterTimeout'
        throw ('Publisher operation {0} exceeded {1} seconds. See {2}.' -f $key, $TimeoutSeconds, $runDirectory)
    }
    if ($null -eq $completion) {
        throw ('Publisher operation {0} exited without a valid completion receipt. See {1}.' -f $key, $stderr)
    }
    if (-not $completion.operationSucceeded -or ($closure.workerTeardown -eq 'Exited' -and $worker.ExitCode -ne 0)) {
        throw ('Publisher operation {0} failed with exit code {1}. See {2}.' -f $key, $worker.ExitCode, $stderr)
    }
    $reportName = 'native-' + $Case.ToLowerInvariant() + '.json'
    if ($Operation -eq 'Validate') {
        $reportName = 'native-validation.json'
    }
    $report = Get-Content -LiteralPath (Join-Path $output $reportName) -Raw | ConvertFrom-Json
    if ($Operation -eq 'Validate' -and $report.stage -ne $Stage) {
        throw 'The completed validation report does not match the requested stage.'
    }
    $succeeded = $true
} catch {
    $primaryError = $_
} finally {
    try {
        if ($null -ne $worker) {
            if (-not $worker.HasExited) {
                Stop-Process -Id $worker.Id -ErrorAction Stop
                if (-not $worker.WaitForExit(10000)) {
                    throw 'The owned native-operation worker did not exit after bounded cleanup.'
                }
                $closure.workerStopped = $true
            }
            $worker.Dispose()
        }
        if (Test-Path -LiteralPath $receiptPath) {
            $receipt = Read-NativePublisherReceipt -Path $receiptPath
            $native = Get-Process -Id $receipt.processId -ErrorAction SilentlyContinue
            if ($null -eq $native) {
                $closure.native = 'Exited'
            } else {
                if ($native.ProcessName -eq 'MSPUB' -and $native.StartTime.ToUniversalTime().Ticks -eq [long]$receipt.startTimeUtcTicks) {
                    if (-not $native.WaitForExit(3000)) {
                        Stop-Process -Id $native.Id -ErrorAction Stop
                        if (-not $native.WaitForExit(10000)) {
                            throw 'The owned Publisher process did not exit after bounded cleanup.'
                        }
                        $closure.native = 'StoppedRecordedProcess'
                    } else {
                        $closure.native = 'Exited'
                    }
                    $native.Dispose()
                } else {
                    $closure.native = 'DifferentIdentityPreserved'
                    throw 'The recorded Publisher PID has a different process identity and is left untouched.'
                }
            }
        } else {
            Write-Warning 'No native process identity was recorded. Any remaining Publisher process is left untouched.'
            if ($succeeded) {
                throw 'An otherwise completed operation has no native process identity receipt.'
            }
        }
    } catch {
        $cleanupError = $_
        $closure.error = $_.Exception.Message
    }
    try {
        $closureBytes = [Text.UTF8Encoding]::new($false).GetBytes(($closure | ConvertTo-Json -Depth 4))
        $closureStream = [IO.FileStream]::new((Join-Path $runDirectory ($key + '.closure.json')), [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        try {
            $closureStream.Write($closureBytes, 0, $closureBytes.Length)
        } finally {
            $closureStream.Dispose()
        }
    } catch {
        if ($null -eq $cleanupError) {
            $cleanupError = $_
        }
    }
}
if ($null -ne $primaryError) {
    if ($null -ne $cleanupError) {
        Write-Warning ('Native cleanup also failed: ' + $cleanupError.Exception.Message)
    }
    throw $primaryError
}
if ($null -ne $cleanupError) {
    throw $cleanupError
}
$report
