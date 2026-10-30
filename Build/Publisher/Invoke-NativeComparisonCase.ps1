<#
.SYNOPSIS
Creates one independent native Publisher comparison case.
.DESCRIPTION
Internal worker for New-NativeComparisonCorpus.ps1. Run the controller instead:
it supplies the deadline and preserves uncertain processes on startup failures.
Corpus content is synthetic, fonts/colors are explicit, and macros stay disabled.
Candidate validation performs only its requested Inspect, Resave or Full stage;
inventories are persisted before PDF exports and repair remains unassessed.
NativeEditText optionally replaces the sole unlinked physical text box before
native save-as during Validate/Full; it never modifies the original input.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputDirectory,

    [ValidateSet('Corpus', 'Validate')]
    [string] $Operation = 'Corpus',

    [string] $InputPath,

    [ValidateSet('Inspect', 'Resave', 'Full')]
    [string] $Stage = 'Full',

    [ValidateLength(1, 4096)]
    [ValidatePattern('\A[^\p{Cc}\p{Zl}\p{Zp}]+\z')]
    [string] $NativeEditText,

    [Parameter(Mandatory)]
    [ValidateSet('Blank', 'BlankPages', 'Text', 'TextLong', 'TextBlank', 'TextParagraphs', 'TextParagraphAttributes', 'TextStories', 'TextFormattingPages', 'TextFormattingMany', 'TextCharacterPages', 'TextCharacterAttributes', 'TextUnderlinePatterns', 'ZeroAnchors', 'Rectangle', 'Image', 'Newsletter', 'TableLabel', 'Brochure', 'MultipleMasters', 'TwoPageMaster', 'PictureOwnership')]
    [string] $Case,

    [Parameter(Mandatory)]
    [string] $ReceiptPath,

    [Parameter(Mandatory)]
    [int] $ControllerProcessId,

    [Parameter(Mandatory)]
    [long] $ControllerStartTimeUtcTicks
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
$nativeEditRequested = $PSBoundParameters.ContainsKey('NativeEditText')
if ($nativeEditRequested -and ($Operation -ne 'Validate' -or $Stage -ne 'Full')) {
    throw 'NativeEditText requires candidate validation with Stage Full.'
}
if ($env:OS -ne 'Windows_NT' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Run the corpus controller on Windows; this worker requires an STA process.'
}
$workerIdentity = Get-CimInstance -ClassName Win32_Process -Filter ('ProcessId=' + $PID) -ErrorAction Stop
$controllerIdentity = Get-Process -Id $ControllerProcessId -ErrorAction Stop
if ($workerIdentity.ParentProcessId -ne $ControllerProcessId -or
    $controllerIdentity.StartTime.ToUniversalTime().Ticks -ne $ControllerStartTimeUtcTicks) {
    throw 'This dedicated worker must be launched directly by its recorded native-operation controller.'
}
$baseline = @(Get-Process -Name MSPUB -ErrorAction SilentlyContinue)
if ($baseline.Count -gt 0) {
    throw 'A Publisher session already exists. No native work was started.'
}
$operationStart = [DateTime]::UtcNow
$output = [IO.Path]::GetFullPath($OutputDirectory)
$stem = 'native-' + $Case.ToLowerInvariant()
$publication = Join-Path $output ($stem + '.pub')
$reference = Join-Path $output ($stem + '.pdf')
$contractPath = Join-Path $output ($stem + '.json')
$diagnosticPrefix = $ReceiptPath.Substring(0, $ReceiptPath.Length - '.process.json'.Length)
$failurePath = $diagnosticPrefix + '.failure.json'
$cleanupPath = $diagnosticPrefix + '.cleanup.json'
$reservedArtifacts = @($publication, $reference, $contractPath)
if ($Operation -eq 'Validate') {
    . (Join-Path $PSScriptRoot 'Get-NativeValidationArtifactNames.ps1')
    $reservedArtifacts = Get-NativeValidationArtifactNames -Stage $Stage |
        ForEach-Object { Join-Path $output $_ }
} elseif ($Case -in @('Image', 'Newsletter', 'TableLabel', 'PictureOwnership')) {
    $reservedArtifacts += Join-Path $output ($stem + '.png')
}
foreach ($path in @($reservedArtifacts) + @($ReceiptPath, $failurePath, $cleanupPath)) {
    if (Test-Path -LiteralPath $path) {
        throw ('Existing evidence is not overwritten: ' + $path)
    }
}

function Write-NewNativeJson {
    param([string] $Path, $Value, [int] $Depth = 14)
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($Value | ConvertTo-Json -Depth $Depth))
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try {
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
    }
}

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class PublisherCorpusNativeWindow {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@

$application = $null
$document = $null
$ownedProcess = $null
$nativeVerified = $false
$primaryFailure = $false
$phase = 'ApplicationStartup'
$objects = [Collections.Generic.List[object]]::new()
$phaseReceiptFailures = [Collections.Generic.List[string]]::new()
$expectedShapes = [Collections.Generic.List[object]]::new()
$expectedText = [Collections.Generic.List[string]]::new()
$width = 612
$height = 792
$pageCount = 1
if ($Case -eq 'TableLabel') {
    $width = 288
    $height = 216
} elseif ($Case -eq 'Brochure') {
    $width = 792
    $height = 612
}
$layout = $null
. (Join-Path $PSScriptRoot 'Add-NativeComparisonLayout.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeComparisonPictures.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeCharacterAttributeControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeParagraphFormatControl.ps1')
. (Join-Path $PSScriptRoot 'Write-NativePublisherPhaseReceipt.ps1')

function Set-NativePhase {
    param([string] $Value)
    $script:phase = $Value
    try {
        if ($nativeVerified) {
            $processReceipt.phase = $Value
            Write-NativePublisherPhaseReceipt -Path $ReceiptPath -Value $processReceipt
        } else {
            $startup = [ordered]@{
                schema = 'officeimo.publisher.native-startup.v1'
                workerProcessId = $PID
                controllerProcessId = $ControllerProcessId
                phase = $Value
                nativeProcessIdentified = $false
            }
            [IO.File]::WriteAllText(($ReceiptPath + '.startup.json'), ($startup | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
        }
    } catch {
        if ($phaseReceiptFailures.Count -eq 0) {
            $phaseReceiptFailures.Add($_.Exception.Message)
        }
    }
}

function Register-NativeWindowIdentity {
    param([switch] $AllowNoWindow)
    try {
        $window = $application.ActiveWindow
    } catch {
        if ($AllowNoWindow) {
            Set-NativePhase ('ActiveWindowUnavailable-' + $_.Exception.HResult)
            return $false
        }
        throw
    }
    if ($null -eq $window -and $AllowNoWindow) {
        Set-NativePhase 'ActiveWindowUnavailable'
        return $false
    }
    $objects.Add($window)
    [uint32]$nativeId = 0
    [void][PublisherCorpusNativeWindow]::GetWindowThreadProcessId([IntPtr]([long]$window.hWnd), [ref]$nativeId)
    $nativeProcess = Get-Process -Id $nativeId -ErrorAction Stop
    if ($nativeProcess.ProcessName -ne 'MSPUB' -or $nativeProcess.StartTime.ToUniversalTime() -lt $operationStart.AddSeconds(-1)) {
        throw 'The COM window does not identify a newly started Publisher process; no application shutdown is attempted.'
    }
    $script:ownedProcess = $nativeProcess
    $script:nativeVerified = $true
    $script:processReceipt = [ordered]@{
        workerProcessId = $PID
        processId = $nativeId
        startTimeUtcTicks = $nativeProcess.StartTime.ToUniversalTime().Ticks.ToString([Globalization.CultureInfo]::InvariantCulture)
        executable = $nativeProcess.Path
        case = $Case
        phase = 'ApplicationIdentified'
        operationSucceeded = $false
    }
    [IO.File]::WriteAllText($ReceiptPath, ($processReceipt | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    $window.Visible = $false
    return $true
}

function Get-NativeShapeInventory {
    # Read the same supported shape fields on physical and master pages, so
    # master markers have independent native geometry and text observations.
    param($Page)
    $nativeShapes = [Collections.Generic.List[object]]::new()
    for ($shapeIndex = 1; $shapeIndex -le $Page.Shapes.Count; $shapeIndex++) {
        $shape = $Page.Shapes.Item($shapeIndex)
        $objects.Add($shape)
        $text = $null
        $font = $null
        $nextLink = $false
        $previousLink = $false
        if ($shape.HasTextFrame -eq -1) {
            $frame = $shape.TextFrame
            $range = $frame.TextRange
            $objects.Add($frame)
            $objects.Add($range)
            $text = [string]$range.Text
            $font = [string]$range.Font.Name
            $nextLink = $frame.HasNextLink -eq -1
            $previousLink = $frame.HasPreviousLink -eq -1
        }
        $nativeShapes.Add([pscustomobject]@{
            name = [string]$shape.Name
            type = [int]$shape.Type
            left = [double]$shape.Left
            top = [double]$shape.Top
            width = [double]$shape.Width
            height = [double]$shape.Height
            text = $text
            font = $font
            nextLink = $nextLink
            previousLink = $previousLink
        })
    }
    return $nativeShapes.ToArray()
}

function Get-NativeInventory {
    param($Publication)
    $nativePages = [Collections.Generic.List[object]]::new()
    for ($pageIndex = 1; $pageIndex -le $Publication.Pages.Count; $pageIndex++) {
        $currentPage = $Publication.Pages.Item($pageIndex)
        $objects.Add($currentPage)
        $masterInfo = $null
        $masterStatus = 'None'
        $masterError = $null
        try {
            $master = $currentPage.Master
            if ($null -ne $master) {
                $objects.Add($master)
                $masterInfo = [pscustomobject]@{ pageId = [int]$master.PageID; name = [string]$master.Name }
                $masterStatus = 'Observed'
            }
        } catch {
            $masterStatus = 'Unavailable'
            $masterError = $_.Exception.HResult
        }
        $nativePages.Add([pscustomobject]@{
            index = $pageIndex
            pageId = [int]$currentPage.PageID
            pageIndex = [int]$currentPage.PageIndex
            pageNumber = [string]$currentPage.PageNumber
            pageType = [int]$currentPage.PageType
            name = [string]$currentPage.Name
            masterStatus = $masterStatus
            master = $masterInfo
            masterErrorHResult = $masterError
            shapes = @(Get-NativeShapeInventory $currentPage)
        })
    }
    $nativeMasters = [Collections.Generic.List[object]]::new()
    for ($masterIndex = 1; $masterIndex -le $Publication.MasterPages.Count; $masterIndex++) {
        $masterPage = $Publication.MasterPages.Item($masterIndex)
        $objects.Add($masterPage)
        # Page.Master is intentionally not read on a master: the documented API
        # rejects it. Native master definitions remain separate from Pages.
        $nativeMasters.Add([pscustomobject]@{
            index = $masterIndex
            pageId = [int]$masterPage.PageID
            pageIndex = [int]$masterPage.PageIndex
            pageType = [int]$masterPage.PageType
            name = [string]$masterPage.Name
            isTwoPageMaster = [bool]$masterPage.IsTwoPageMaster
            shapeCount = [int]$masterPage.Shapes.Count
            shapes = @(Get-NativeShapeInventory $masterPage)
        })
    }
    return [pscustomobject]@{
        pageCount = [int]$Publication.Pages.Count
        widthPoints = [double]$Publication.PageSetup.PageWidth
        heightPoints = [double]$Publication.PageSetup.PageHeight
        readOnly = [bool]$Publication.ReadOnly
        pages = $nativePages.ToArray()
        masterPages = $nativeMasters.ToArray()
    }
}

function Add-TextFrame {
    param($Page, [string] $Name, [string] $Text, [double] $Left, [double] $Top, [double] $Width, [double] $Height, [double] $Size = 12)
    $shape = $Page.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.Name = $Name
    $shape.Line.Visible = 0
    $shape.Fill.Visible = 0
    $frame = $shape.TextFrame
    $frame.MarginLeft = 0
    $frame.MarginRight = 0
    $frame.MarginTop = 0
    $frame.MarginBottom = 0
    $range = $frame.TextRange
    $range.Text = $Text
    $range.Font.Name = 'Arial'
    $range.Font.Size = $Size
    $range.Font.Color.RGB = 0
    $objects.Add($shape)
    $objects.Add($frame)
    $objects.Add($range)
    if ($Text.Length -gt 0) {
        $expectedText.Add($Text)
    }
    $expectedShapes.Add([pscustomobject]@{ name = $Name; kind = 'TextFrame'; left = $Left; top = $Top; width = $Width; height = $Height; font = 'Arial'; size = $Size })
    return $shape
}

function New-ReferencePicture {
    param([string] $Path)
    Add-Type -AssemblyName System.Drawing
    $bitmap = [Drawing.Bitmap]::new(32, 32)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([Drawing.Color]::White)
        $graphics.FillRectangle([Drawing.Brushes]::Red, 0, 0, 16, 16)
        $graphics.FillRectangle([Drawing.Brushes]::Blue, 16, 0, 16, 16)
        $graphics.FillRectangle([Drawing.Brushes]::Green, 0, 16, 16, 16)
        $imageStream = [IO.FileStream]::new($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        try {
            $bitmap.Save($imageStream, [Drawing.Imaging.ImageFormat]::Png)
        } finally {
            $imageStream.Dispose()
        }
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

try {
    Set-NativePhase 'ActivateApplication'
    $application = New-Object -ComObject Publisher.Application
    Set-NativePhase 'SetAutomationSecurity'
    $application.AutomationSecurity = 3
    Set-NativePhase 'ObserveWindowBeforeDocument'
    $identifiedBeforeDocument = Register-NativeWindowIdentity -AllowNoWindow
    Set-NativePhase 'CreateDocument'
    $document = $application.NewDocument()
    if (-not $identifiedBeforeDocument) {
        Set-NativePhase 'ObserveWindowAfterDocument'
        $null = Register-NativeWindowIdentity
    }
    if ($Operation -eq 'Validate') {
        Set-NativePhase 'CaptureCandidate'
        $candidate = [IO.Path]::GetFullPath($InputPath)
        $candidateHash = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToLowerInvariant()
        $snapshot = Join-Path $output 'candidate-input.pub'
        $initialInventoryPath = Join-Path $output 'native-initial-inventory.json'
        $initialPdf = Join-Path $output 'native-initial.pdf'
        $resaved = Join-Path $output 'native-resaved.pub'
        $resavedInventoryPath = Join-Path $output 'native-resaved-inventory.json'
        $resavedPdf = Join-Path $output 'native-resaved.pdf'
        $validationPath = Join-Path $output 'native-validation.json'
        foreach ($path in $reservedArtifacts) {
            if (Test-Path -LiteralPath $path) {
                throw ('Existing validation evidence is not overwritten: ' + $path)
            }
        }
        $qualifiedOperations = [Collections.Generic.List[string]]::new()
        $repairStatus = 'Unassessed: COM Open success does not expose whether native repair occurred.'
        [IO.File]::Copy($candidate, $snapshot, $false)
        if ((Get-FileHash -LiteralPath $snapshot -Algorithm SHA256).Hash.ToLowerInvariant() -ne $candidateHash) {
            throw 'The candidate changed while taking the qualification snapshot.'
        }
        $document.Close()
        $document = $null
        Set-NativePhase 'OpenCandidateReadOnly'
        $document = $application.Open($snapshot, $true, $false)
        Set-NativePhase 'InitialInventory'
        $initialInventory = Get-NativeInventory $document
        if (-not $initialInventory.readOnly) {
            throw 'Publisher did not open the initial candidate read-only.'
        }
        $qualifiedOperations.Add('ReadOnlyOpen')
        $qualifiedOperations.Add('InitialInventory')
        $producer = [ordered]@{
            name = 'Microsoft Publisher'
            version = [string]$application.Version
            executableVersion = (Get-Item -LiteralPath $ownedProcess.Path).VersionInfo.FileVersion
            automationSecurity = 3
        }
        # Persist the observed input before an export can crash the application.
        # This is an observation receipt, not completion of the requested stage.
        Write-NewNativeJson -Path $initialInventoryPath -Value ([ordered]@{
            schema = 'officeimo.publisher.native-inventory.v1'
            stage = $Stage
            observation = 'Initial'
            status = 'Observed'
            candidateSha256 = $candidateHash
            producer = $producer
            qualifiedOperations = $qualifiedOperations.ToArray()
            repairStatus = $repairStatus
            inventory = $initialInventory
        })
        if ($Stage -eq 'Full') {
            Set-NativePhase 'InitialPdfExport'
            $document.ExportAsFixedFormat(2, $initialPdf)
            $qualifiedOperations.Add('InitialPdfExport')
        }
        Set-NativePhase 'CloseInitialDocument'
        $document.Close()
        $document = $null
        if ((Get-FileHash -LiteralPath $snapshot -Algorithm SHA256).Hash.ToLowerInvariant() -ne $candidateHash) {
            throw 'The initial read-only observation changed the input snapshot.'
        }
        $resavedInventory = $null
        $nativeTextMutation = $null
        if ($Stage -ne 'Inspect') {
            # Only this task-owned snapshot is opened writable. The original
            # candidate remains outside the native application's mutation path.
            Set-NativePhase 'OpenSnapshotWritable'
            $document = $application.Open($snapshot, $false, $false)
            if ($nativeEditRequested) {
                Set-NativePhase 'NativeTextMutation'
                if ($document.ReadOnly -or $document.Pages.Count -ne 1) {
                    throw 'Native text mutation requires one writable physical page.'
                }
                $editPage = $document.Pages.Item(1)
                $objects.Add($editPage)
                if ($editPage.Shapes.Count -ne 1) {
                    throw 'Native text mutation requires exactly one physical text box.'
                }
                $editShape = $editPage.Shapes.Item(1)
                $objects.Add($editShape)
                if ($editShape.Type -ne 17 -or $editShape.HasTextFrame -ne -1) {
                    throw 'Native text mutation requires an ordinary text box.'
                }
                $editFrame = $editShape.TextFrame
                $objects.Add($editFrame)
                if ($editFrame.HasNextLink -eq -1 -or $editFrame.HasPreviousLink -eq -1) {
                    throw 'Native text mutation does not accept linked frames.'
                }
                $editRange = $editFrame.TextRange
                $objects.Add($editRange)
                $beforeText = [string]$editRange.Text
                $editRange.Text = $NativeEditText
                $observedText = [string]$editRange.Text
                # A freshly assigned range can omit the final paragraph CR
                # that Publisher materializes when saving the publication.
                if ($observedText -cne $NativeEditText -and $observedText -cne ($NativeEditText + "`r")) {
                    throw 'Publisher did not retain the requested native text mutation.'
                }
                $nativeTextMutation = [ordered]@{
                    shapeName = [string]$editShape.Name
                    before = $beforeText
                    requested = $NativeEditText
                    observed = $observedText
                }
                $qualifiedOperations.Add('NativeTextMutation')
            }
            Set-NativePhase 'NativeResave'
            # 1 is pbFilePublication; exclude task output from recent files.
            $document.SaveAs($resaved, 1, $false)
            $qualifiedOperations.Add('NativeSaveAs')
            Set-NativePhase 'CloseNativeResave'
            $document.Close()
            $document = $null
            Set-NativePhase 'ReopenNativeResave'
            $document = $application.Open($resaved, $true, $false)
            Set-NativePhase 'ResavedInventory'
            $resavedInventory = Get-NativeInventory $document
            if (-not $resavedInventory.readOnly) {
                throw 'Publisher did not reopen the native resave read-only.'
            }
            if ($nativeEditRequested -and
                ($resavedInventory.pageCount -ne 1 -or
                 $resavedInventory.pages[0].shapes.Count -ne 1 -or
                 $resavedInventory.pages[0].shapes[0].text -cne ($NativeEditText + "`r"))) {
                throw 'The native text mutation did not survive save-as and reopen.'
            }
            $qualifiedOperations.Add('ReadOnlyReopen')
            $qualifiedOperations.Add('ResavedInventory')
            Write-NewNativeJson -Path $resavedInventoryPath -Value ([ordered]@{
                schema = 'officeimo.publisher.native-inventory.v1'
                stage = $Stage
                observation = 'Resaved'
                status = 'Observed'
                candidateSha256 = $candidateHash
                producer = $producer
                qualifiedOperations = $qualifiedOperations.ToArray()
                repairStatus = $repairStatus
                inventory = $resavedInventory
            })
            if ($Stage -eq 'Full') {
                Set-NativePhase 'ResavedPdfExport'
                $document.ExportAsFixedFormat(2, $resavedPdf)
                $qualifiedOperations.Add('ResavedPdfExport')
            }
            Set-NativePhase 'CloseReopenedResave'
            $document.Close()
            $document = $null
        }
        if ((Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToLowerInvariant() -ne $candidateHash) {
            throw 'The original candidate changed during qualification.'
        }
        $files = [Collections.Generic.List[object]]::new()
        foreach ($path in $reservedArtifacts) {
            if ($path -eq $validationPath) {
                continue
            }
            if (-not (Test-Path -LiteralPath $path) -or (Get-Item -LiteralPath $path).Length -eq 0) {
                throw ('Native qualification output is missing or empty: ' + $path)
            }
            $files.Add([pscustomobject]@{ name = [IO.Path]::GetFileName($path); sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() })
        }
        $status = 'OpenedExportedResavedAndReopened'
        if ($Stage -eq 'Inspect') {
            $status = 'OpenedReadOnlyAndInventoried'
        } elseif ($Stage -eq 'Resave') {
            $status = 'OpenedResavedAndReopened'
        }
        $report = [ordered]@{
            schema = 'officeimo.publisher.native-validation.v1'
            stage = $Stage
            status = $status
            stageCompleted = $true
            qualifiedOperations = $qualifiedOperations.ToArray()
            qualificationScope = 'Only the listed native operations are established. Native repair and rendering fidelity remain unassessed.'
            candidateSha256 = $candidateHash
            sourceUnchanged = $true
            producer = $producer
            repairStatus = $repairStatus
            initial = $initialInventory
            resaved = $resavedInventory
            files = $files.ToArray()
        }
        if ($nativeEditRequested) {
            $report.nativeTextMutation = $nativeTextMutation
        }
        Write-NewNativeJson -Path $validationPath -Value $report
        Set-NativePhase 'Shutdown'
        return
    }
    $phase = 'AuthorCorpusCase'
    $document.PageSetup.PageWidth = $width
    $document.PageSetup.PageHeight = $height
    $page = $document.Pages.Item(1)
    $objects.Add($page)
    switch ($Case) {
        'Blank' { }
        'BlankPages' {
            $pageCount = 2
            $second = $document.Pages.Add(1, 1)
            $objects.Add($second)
        }
        'Text' {
            $null = Add-TextFrame $page 'MinimalText' 'Publisher text' 72 72 360 72
        }
        'TextLong' {
            $null = Add-TextFrame $page 'MinimalText' 'Publisher text with a longer replacement' 72 72 360 72
        }
        'TextParagraphs' {
            $null = Add-TextFrame $page 'ParagraphText' "First paragraph`r`rLast paragraph" 72 72 360 144
        }
        'TextParagraphAttributes' {
            $pageCount = Add-NativeParagraphFormatControl -Publication $document -FirstPage $page
        }
        'TextStories' {
            $null = Add-TextFrame $page 'FirstStory' 'First independent story' 72 72 360 72
            $null = Add-TextFrame $page 'MiddleStory' "Middle paragraph`r`rLast middle paragraph" 72 216 360 108
            $null = Add-TextFrame $page 'LastStory' 'Last independent story' 72 396 360 72
        }
        'TextFormattingPages' {
            $paragraphs = @(1..96 | ForEach-Object { 'Paragraph {0:D2} preserves its native formatting.' -f $_ })
            $null = Add-TextFrame $page 'FormattingPages' ($paragraphs -join "`r") 72 72 360 648
        }
        'TextFormattingMany' {
            $paragraphs = @(1..240 | ForEach-Object { 'Paragraph {0:D3} preserves its native formatting.' -f $_ })
            $null = Add-TextFrame $page 'ManyFormattingPages' ($paragraphs -join "`r") 72 72 360 648
        }
        'TextCharacterPages' {
            $null = Add-TextFrame $page 'Character12' 'Character page font twelve' 72 72 360 72 12
            $null = Add-TextFrame $page 'Character14' 'Character page font fourteen' 72 216 360 72 14
            $null = Add-TextFrame $page 'Character16' 'Character page font sixteen' 72 396 360 72 16
            $null = Add-TextFrame $page 'Character18' 'Character page font eighteen' 72 576 360 72 18
        }
        'TextCharacterAttributes' {
            $pageCount = Add-NativeCharacterAttributeControl -Publication $document -FirstPage $page
        }
        'TextUnderlinePatterns' {
            $pageCount = Add-NativeCharacterAttributeControl -Publication $document -FirstPage $page -UnderlinePatterns
        }
        'ZeroAnchors' {
            $null = Add-TextFrame $page 'ZeroLeft' 'Left at native zero' 306 72 234 72
            $null = Add-TextFrame $page 'ZeroTop' 'Top at native zero' 72 396 360 72
            $null = Add-TextFrame $page 'ZeroRight' 'Right at native zero' 72 216 234 72
            $null = Add-TextFrame $page 'ZeroBottom' 'Bottom at native zero' 72 288 360 108
        }
        'TextBlank' {
            $pageCount = 2
            $null = Add-TextFrame $page 'MinimalText' 'Publisher text' 72 72 360 72
            $second = $document.Pages.Add(1, 1)
            $objects.Add($second)
        }
        'Rectangle' {
            $rectangle = $page.Shapes.AddShape(1, 72, 72, 144, 72)
            $rectangle.Name = 'MinimalRectangle'
            $rectangle.Fill.ForeColor.RGB = 255
            $rectangle.Line.ForeColor.RGB = 0
            $rectangle.Line.Weight = 1
            $objects.Add($rectangle)
            $expectedShapes.Add([pscustomobject]@{ name = 'MinimalRectangle'; kind = 'Rectangle'; left = 72; top = 72; width = 144; height = 72 })
        }
        'Image' {
            $picturePath = Join-Path $output ($stem + '.png')
            New-ReferencePicture $picturePath
            $image = $page.Shapes.AddPicture($picturePath, 0, -1, 72, 72, 96, 96)
            $image.Name = 'MinimalImage'
            $objects.Add($image)
            $expectedShapes.Add([pscustomobject]@{ name = 'MinimalImage'; kind = 'Image'; left = 72; top = 72; width = 96; height = 96 })
        }
        'Newsletter' {
            $pageCount = 2
            $second = $document.Pages.Add(1, 1)
            $objects.Add($second)
            $title = Add-TextFrame $page 'NewsletterTitle' 'Publisher comparison newsletter' 54 36 504 42 20
            $title.TextFrame.TextRange.Font.Bold = -1
            $paragraph = 'Linked stories preserve reading order across publication pages. This paragraph uses explicit Arial text, ordinary spaces, and a visible page continuation.'
            $story = ((1..10 | ForEach-Object { ('Section {0}. {1}' -f $_, $paragraph) }) -join "`r")
            $first = Add-TextFrame $page 'StoryStart' $story 54 108 228 120 11
            $last = Add-TextFrame $second 'StoryEnd' '' 54 72 468 600 11
            $first.TextFrame.NextLinkedTextFrame = $last.TextFrame
            $picturePath = Join-Path $output ($stem + '.png')
            New-ReferencePicture $picturePath
            $image = $page.Shapes.AddPicture($picturePath, 0, -1, 330, 108, 144, 144)
            $image.Name = 'NewsletterImage'
            $objects.Add($image)
            $expectedShapes.Add([pscustomobject]@{ name = 'NewsletterImage'; kind = 'Image'; left = 330; top = 108; width = 144; height = 144 })
            if ($first.TextFrame.HasNextLink -ne -1 -or $last.TextFrame.HasPreviousLink -ne -1) {
                throw 'Publisher did not establish reciprocal newsletter frame links.'
            }
        }
        'TableLabel' {
            # Unicode escapes keep the maintained script parseable with PS5.1's default file encoding.
            $multilingual = 'Za' + [char]0x017C + [char]0x00F3 + [char]0x0142 + [char]0x0107 + ' g' + [char]0x0119 + [char]0x015B + 'l' + [char]0x0105 + ' ja' + [char]0x017A + [char]0x0144 + ' / ' + [char]0x0391 + [char]0x03BB + [char]0x03C6 + [char]0x03B1
            $null = Add-TextFrame $page 'LabelTitle' $multilingual 18 12 252 42 12
            $table = $page.Shapes.AddTable(2, 2, 18, 66, 180, 72)
            $table.Name = 'LabelTable'
            $objects.Add($table)
            $values = @('Item', 'Count', 'Alpha', '2')
            for ($row = 1; $row -le 2; $row++) {
                for ($column = 1; $column -le 2; $column++) {
                    $cell = $table.Table.Cells($row, $column, $row, $column).Item(1)
                    $objects.Add($cell)
                    $cell.TextRange.Text = $values[($row - 1) * 2 + $column - 1]
                    $cell.TextRange.Font.Name = 'Arial'
                    $cell.TextRange.Font.Size = 10
                    $cell.TextRange.Font.Color.RGB = 0
                    if ($row -eq 1) {
                        $cell.Fill.ForeColor.RGB = 15790320
                        $cell.TextRange.Font.Bold = -1
                    } else {
                        $cell.Fill.ForeColor.RGB = 16777215
                    }
                }
            }
            $expectedShapes.Add([pscustomobject]@{ name = 'LabelTable'; kind = 'Table'; rows = 2; columns = 2; cells = $values; left = 18; top = 66; width = 180; height = 72 })
            foreach ($value in $values) {
                $expectedText.Add($value)
            }
            $picturePath = Join-Path $output ($stem + '.png')
            New-ReferencePicture $picturePath
            $image = $page.Shapes.AddPicture($picturePath, 0, -1, 216, 72, 54, 54)
            $image.Name = 'LabelImage'
            $objects.Add($image)
            $expectedShapes.Add([pscustomobject]@{ name = 'LabelImage'; kind = 'Image'; left = 216; top = 72; width = 54; height = 54 })
        }
        { $_ -in @('Brochure', 'MultipleMasters', 'TwoPageMaster') } {
            $layout = Add-NativeComparisonLayout -Case $Case -Document $document -FirstPage $page
            $pageCount = $layout.pages
        }
        'PictureOwnership' {
            $picturePath = Join-Path $output ($stem + '.png')
            New-ReferencePicture $picturePath
            $layout = Add-NativeComparisonPictures -Document $document -FirstPage $page -PicturePath $picturePath
            $pageCount = $layout.pages
        }
    }
    Set-NativePhase 'ObserveCorpusPageCount'
    if ($document.Pages.Count -ne $pageCount) {
        throw ('Native page count differs from the declared contract: ' + $document.Pages.Count)
    }
    Set-NativePhase 'InventoryCorpusCase'
    $inventory = Get-NativeInventory $document
    if ($null -ne $layout) {
        foreach ($assignment in $layout.masterAssignments) {
            $observed = $inventory.pages[$assignment.pageIndex - 1]
            if ($observed.masterStatus -ne 'Observed' -or $observed.master.pageId -ne $assignment.masterPageId) {
                throw ('Publisher did not retain the declared master assignment for physical page ' + $assignment.pageIndex)
            }
        }
    }
    Set-NativePhase 'SaveCorpusCase'
    # 1 is pbFilePublication; exclude task output from recent files.
    $document.SaveAs($publication, 1, $false)
    Set-NativePhase 'ExportCorpusReference'
    $document.ExportAsFixedFormat(2, $reference)
    Set-NativePhase 'ObserveCorpusProducer'
    $producerVersion = [string]$application.Version
    $executableVersion = (Get-Item -LiteralPath $ownedProcess.Path).VersionInfo.FileVersion
    # Publisher holds an exclusive stream handle until Close. Capture native
    # observations first, then close before hashing the exact output bytes.
    Set-NativePhase 'CloseCorpusCase'
    $document.Close()
    $document = $null
    foreach ($path in @($publication, $reference)) {
        if (-not (Test-Path -LiteralPath $path) -or (Get-Item -LiteralPath $path).Length -eq 0) {
            throw ('Native output is missing or empty: ' + $path)
        }
    }
    $fontPath = Join-Path $env:WINDIR 'Fonts/arial.ttf'
    $font = [ordered]@{ family = 'Arial'; file = 'arial.ttf'; sha256 = $null }
    if (Test-Path -LiteralPath $fontPath) {
        $font.sha256 = (Get-FileHash -LiteralPath $fontPath -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    $contract = [ordered]@{
        schema = 'officeimo.publisher.native-case.v1'
        case = $Case
        status = 'Produced'
        license = 'MIT'
        producer = [ordered]@{
            name = 'Microsoft Publisher'
            version = $producerVersion
            executableVersion = $executableVersion
            powershell = $PSVersionTable.PSVersion.ToString()
            culture = [Globalization.CultureInfo]::CurrentCulture.Name
            textFont = $font
            appearance = 'Explicit frame fonts, margins and object colors; untouched native publication defaults.'
            automationSecurity = 3
        }
        expected = [ordered]@{ pages = $pageCount; widthPoints = $width; heightPoints = $height; text = $expectedText.ToArray(); shapes = $expectedShapes.ToArray() }
        nativePages = $inventory.pages
        nativeMasterPages = $inventory.masterPages
        files = [ordered]@{
            publication = [ordered]@{ name = [IO.Path]::GetFileName($publication); sha256 = (Get-FileHash -LiteralPath $publication -Algorithm SHA256).Hash.ToLowerInvariant() }
            reference = [ordered]@{ name = [IO.Path]::GetFileName($reference); sha256 = (Get-FileHash -LiteralPath $reference -Algorithm SHA256).Hash.ToLowerInvariant() }
        }
    }
    if ($null -ne $layout) {
        $contract.expected.masterAssignments = $layout.masterAssignments
        $contract.expected.panelWidthPoints = $layout.panelWidthPoints
    }
    if ($Case -eq 'PictureOwnership') {
        $contract.nativePictureObservations = $layout.pictureObservations
        $contract.expected.pictureDuplication = [pscustomobject]@{
            sourceName = 'PhysicalPictureOriginal'
            duplicateName = 'PhysicalPictureDuplicate'
            method = 'Shape.Duplicate'
            nativeImageStoreSharing = 'Unassessed'
        }
    }
    $sourceImage = Join-Path $output ($stem + '.png')
    if (Test-Path -LiteralPath $sourceImage) {
        $contract.files.sourceImage = [ordered]@{ name = [IO.Path]::GetFileName($sourceImage); sha256 = (Get-FileHash -LiteralPath $sourceImage -Algorithm SHA256).Hash.ToLowerInvariant() }
    }
    Write-NewNativeJson -Path $contractPath -Value $contract
    $phase = 'Shutdown'
} catch {
    $primaryFailure = $true
    $failure = [ordered]@{
        schema = 'officeimo.publisher.native-failure.v1'
        operation = $Operation
        case = $Case
        phase = $phase
        message = $_.Exception.Message
        hresult = $_.Exception.HResult
        exceptionType = $_.Exception.GetType().FullName
        errorId = $_.FullyQualifiedErrorId
        position = $_.InvocationInfo.PositionMessage
        scriptStackTrace = $_.ScriptStackTrace
        nativeProcessIdentified = $nativeVerified
    }
    if ($Operation -eq 'Validate') {
        $failure.stage = $Stage
    }
    Write-NewNativeJson -Path $failurePath -Value $failure -Depth 5
    [Console]::Error.WriteLine(('Native {0} failed during {1}: {2}' -f $Operation, $phase, $_.Exception.Message))
} finally {
    $cleanupFailures = [Collections.Generic.List[string]]::new()
    if ($null -ne $document -and $nativeVerified) {
        try {
            Set-NativePhase 'CloseDocument'
            $document.Close()
        } catch {
            $cleanupFailures.Add('Close: ' + $_.Exception.Message)
        }
    }
    # RCW finalization can block Publisher's STA even after successful Close/Quit.
    # The dedicated child exits after durable evidence writes; the controller
    # independently verifies and closes its exact recorded native process.
    if ($null -ne $application -and $nativeVerified) {
        try {
            Set-NativePhase 'QuitApplication'
            $application.Quit()
        } catch {
            $cleanupFailures.Add('Quit: ' + $_.Exception.Message)
        }
    }
    if ($null -ne $ownedProcess) {
        $ownedProcess.Dispose()
    }
    foreach ($receiptFailure in $phaseReceiptFailures) {
        $cleanupFailures.Add('Phase receipt: ' + $receiptFailure)
    }
    if ($cleanupFailures.Count -gt 0) {
        $cleanupReport = [ordered]@{ schema = 'officeimo.publisher.native-cleanup.v1'; primaryFailure = $primaryFailure; failures = $cleanupFailures.ToArray() }
        Write-NewNativeJson -Path $cleanupPath -Value $cleanupReport -Depth 4
        [Console]::Error.WriteLine(('Native cleanup also reported: ' + ($cleanupFailures -join '; ')))
    }
    $exitCode = 0
    if ($primaryFailure -or $cleanupFailures.Count -gt 0) {
        $exitCode = 1
    }
    if ($nativeVerified) {
        $processReceipt.operationSucceeded = $exitCode -eq 0
        Set-NativePhase 'ShutdownComplete'
    }
    # The parent PID and start time were verified against Win32_Process before
    # COM activation. This exit is confined to the controller-created worker.
    [Environment]::Exit($exitCode)
}
