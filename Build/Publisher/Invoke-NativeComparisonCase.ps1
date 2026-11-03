<#
.SYNOPSIS
Creates one independent native Publisher comparison case.
.DESCRIPTION
Internal worker for New-NativeComparisonCorpus.ps1. Run the controller instead:
it supplies the deadline and preserves uncertain processes on startup failures.
Corpus content is synthetic, fonts/colors are explicit, and macros stay disabled.
Candidate validation performs only its requested Inspect, Export, Resave or Full stage;
inventories are persisted before PDF exports and repair remains unassessed.
NativeEditText optionally replaces the sole unlinked physical text box or the
sole table's first unmerged cell or first horizontal 1x2 merged anchor before
native save-as during Validate/Full; it never modifies the original input.
LinkedStoryFirstCharacter instead changes one printable ASCII character at the
head of the qualified three-frame ordinary story without replacing its paragraphs.
NativeAddClusterTable adds a fixed two-by-four table beside the retained
eighteen-by-eighteen table in the writable snapshot during Validate/Full.
.PARAMETER NativeEditTarget
TextBox selects the sole unlinked text box. FirstTableCell requires an unmerged
first cell; FirstMergedTableCell requires the native row/column 1,1 anchor to
span exactly two columns and one row. Table targets require the sole physical
table on one page. All targets require NativeEditText and Validate/Full.
LinkedStoryFirstCharacter requires the sole page 1/3/2 three-frame story and a
single printable ASCII replacement character distinct from its original M.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputDirectory,

    [ValidateSet('Corpus', 'Validate')]
    [string] $Operation = 'Corpus',

    [string] $InputPath,

    [ValidateSet('Inspect', 'Export', 'Resave', 'Full')]
    [string] $Stage = 'Full',

    [AllowEmptyString()]
    [ValidateLength(0, 4096)]
    [ValidatePattern('\A[^\p{Cc}\p{Zl}\p{Zp}]*\z')]
    [string] $NativeEditText,

    [ValidateSet('TextBox', 'FirstTableCell', 'FirstMergedTableCell', 'LinkedStoryFirstCharacter')]
    [string] $NativeEditTarget = 'TextBox',

    [switch] $NativeAddClusterTable,

    [Parameter(Mandatory)]
    [ValidateSet('Blank', 'BlankPages', 'TextPages', 'ObjectPages', 'ObjectPagesEdited', 'MixedObjectPages', 'MixedObjectPagesEdited', 'WrapObjectOrderDefaults', 'WrapObjectOrderLayout', 'WrapTextFrameOutside', 'WrapObjectOrderLeft', 'WrapObjectOrderRight', 'WrapObjectOrderLarger', 'WrapObjectOrderLargerLeft', 'WrapObjectOrderLargerTie', 'WrapObjectOrderLargerColumns', 'WrapObjectOrderOutside', 'WrapObjectOrderRectangleOnly', 'WrapObjectOrderPictureOnly', 'WrapObjectOrderSwap', 'WrapObjectOrder', 'WrapObjectOrderReversed', 'WrapObjectOrderNone', 'MixedObjectOrder', 'MixedObjectOrderReversed', 'MasterContent', 'MasterContentEdited', 'LinkedStory', 'LinkedStoryOverflow', 'MultiParagraphCreation', 'MultiParagraphCreationLinked', 'TableCreation', 'TableCreationEdited', 'TableCreationMerged', 'TableCreationMergedEdited', 'TableMergedDecoration', 'TableMergedDecorationEdited', 'TableMergedDecorationText', 'TableMergedDecorationComposed', 'TableMergedFills', 'TableMergedBorders', 'CaptionTypingPreRange', 'CaptionTypingFullRange', 'CaptionTypingExplicit', 'CaptionTypingPreRangeRefilled', 'CaptionTypingFullRangeRefilled', 'CaptionTypingExplicitRefilled', 'TableDecoration', 'TableDecorationEdited', 'TableDecorationText', 'TableBordersOnly', 'TableBordersOnlyText', 'TableFillsOnly', 'TableBordersOnlyCleared', 'TableUtilityPair', 'TableUtilityPairStyled', 'TableUtilityPairBorders', 'TableUtilityPairFills', 'TableUtilityPairComposed', 'TableUtilityPairSingleFill', 'TableUtilityPairTargetComposed', 'TableClusterNear', 'TableClusterNearEdited', 'TableClusterWide', 'TableFillVisibility', 'TableFillVisibilityRestored', 'TableFillVisibilityText', 'TableBorderTopology', 'TableBordersCleared', 'TableBorderTopologyTextFill', 'TableBorderDefaultFill', 'FrameCreation', 'FrameCreationEdited', 'FrameCreationFrench', 'ImagePages', 'ImagePagesEdited', 'ImagePagesReplacement', 'ImagePagesJpegReplacement', 'ImagePagesJpegCropped', 'ImagePagesCropped', 'SharedImagePages', 'SharedImagePagesReplacement', 'Text', 'TextLong', 'TextBlank', 'TextParagraphs', 'TextNonuniformParagraphs', 'TextNonuniformTextReplacements', 'TextMixedRunTextReplacements', 'TextFormattingTopology', 'TextEmptyRefill', 'TextRunSegmentation', 'TextNonuniformEmptyParagraphs', 'TextUniformEmptyParagraphs', 'TextParagraphAttributes', 'TextAlignedTabs', 'TextFrameLayout', 'TextVerticalPlacement', 'TextFontMetrics', 'TextFallbackAlignment', 'TextFrameBreaks', 'TextLabelFitting', 'TextHyphenation', 'TextStories', 'TextFormattingPages', 'TextFormattingMany', 'TextCharacterPages', 'TextCharacterAttributes', 'TextUnderlinePatterns', 'ZeroAnchors', 'Rectangle', 'Image', 'Newsletter', 'TableLabel', 'TableLabelWrapped', 'TableLabelShortParagraphs', 'TableCellTyping', 'TableCellTypingEarlyRange', 'TableCellParagraphs', 'TableNonuniformParagraphs', 'TableNonuniformTextReplacements', 'TableMixedRunTextReplacements', 'TableFormattingTopology', 'TableEmptyRefill', 'TableRunSegmentation', 'TableNonuniformEmptyParagraphs', 'TableUniformEmptyParagraphs', 'Brochure', 'MultipleMasters', 'TwoPageMaster', 'MultipleMastersEdited', 'TwoPageMasterEdited', 'MultipleMastersReassigned', 'MultipleMastersIgnored', 'TwoPageMasterReassigned', 'TwoPageMasterIgnored', 'PictureOwnership', 'PictureReplacement', 'PictureCropped', 'PictureCropFractions', 'PictureCropFractionsReplacement')]
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
if ($PSBoundParameters.ContainsKey('NativeEditTarget') -and -not $nativeEditRequested) {
    throw 'NativeEditTarget requires NativeEditText.'
}
if ($nativeEditRequested -and ($Operation -ne 'Validate' -or $Stage -ne 'Full')) {
    throw 'NativeEditText requires candidate validation with Stage Full.'
}
if ($NativeAddClusterTable -and ($Operation -ne 'Validate' -or $Stage -ne 'Full')) {
    throw 'NativeAddClusterTable requires candidate validation with Stage Full.'
}
if ($NativeAddClusterTable -and $nativeEditRequested) {
    throw 'NativeAddClusterTable cannot be combined with NativeEditText.'
}
if ($env:OS -ne 'Windows_NT' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Run the corpus controller on Windows; this worker requires an STA process.'
}
if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw 'Native COM operations require the PowerShell 7 worker selected by the controller.'
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
} elseif ($Case -in @('WrapObjectOrderDefaults', 'WrapObjectOrderLayout', 'WrapTextFrameOutside', 'WrapObjectOrderLeft', 'WrapObjectOrderRight', 'WrapObjectOrderLarger', 'WrapObjectOrderLargerLeft', 'WrapObjectOrderLargerTie', 'WrapObjectOrderLargerColumns', 'WrapObjectOrderOutside', 'WrapObjectOrderRectangleOnly', 'WrapObjectOrderPictureOnly', 'WrapObjectOrderSwap', 'WrapObjectOrder', 'WrapObjectOrderReversed', 'WrapObjectOrderNone', 'Image', 'ObjectPages', 'ObjectPagesEdited', 'MixedObjectPages', 'MixedObjectPagesEdited', 'MixedObjectOrder', 'MixedObjectOrderReversed', 'MasterContent', 'MasterContentEdited', 'ImagePages', 'ImagePagesEdited', 'ImagePagesReplacement', 'ImagePagesJpegReplacement', 'ImagePagesJpegCropped', 'ImagePagesCropped', 'SharedImagePages', 'SharedImagePagesReplacement', 'Newsletter', 'TableLabel', 'TableLabelWrapped', 'TableLabelShortParagraphs', 'PictureOwnership', 'PictureReplacement', 'PictureCropped', 'PictureCropFractions', 'PictureCropFractionsReplacement')) {
    $reservedArtifacts += Join-Path $output ($stem + '.png')
    if ($Case -in @('MasterContentEdited', 'PictureReplacement', 'PictureCropFractionsReplacement', 'ImagePages', 'ImagePagesEdited', 'ImagePagesReplacement', 'ImagePagesJpegReplacement', 'ImagePagesJpegCropped', 'ImagePagesCropped', 'SharedImagePages', 'SharedImagePagesReplacement')) {
        $reservedArtifacts += Join-Path $output ($stem + '-replacement.png')
        if ($Case -in @('ImagePagesJpegReplacement', 'ImagePagesJpegCropped')) { $reservedArtifacts += Join-Path $output ($stem + '-replacement.jpg') }
    }
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
$phaseSequence = 0
$objects = [Collections.Generic.List[object]]::new()
$phaseReceiptFailures = [Collections.Generic.List[string]]::new()
$expectedShapes = [Collections.Generic.List[object]]::new()
$expectedText = [Collections.Generic.List[string]]::new()
$width = 612
$height = 792
$pageCount = 1
if ($Case -in @('TableLabel', 'TableLabelWrapped', 'TableLabelShortParagraphs')) {
    $width = 288
    $height = 216
} elseif ($Case -eq 'Brochure') {
    $width = 792
    $height = 612
}
$layout = $null
$captionRangeObservations = $null
$multiParagraphObservations = $null
. (Join-Path $PSScriptRoot 'Add-NativeComparisonLayout.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeComparisonPictures.ps1')
. (Join-Path $PSScriptRoot 'Get-NativePictureCropBasis.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeObjectPagesControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeMixedPagesControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeWrapOrderControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeMasterContentControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeLinkedStoryControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeMultiParagraphControl.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeLinkedStoryEditState.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeTableCreationControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeTableDecorationControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeMergedTableDecorationControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeCaptionTypingControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeTableUtilityPairControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeTableClusterControl.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeTableCellStyle.ps1')
. (Join-Path $PSScriptRoot 'Set-NativeCandidateText.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeCandidateClusterTable.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeFrameCreationControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeImagePagesControl.ps1')
. (Join-Path $PSScriptRoot 'New-NativeReplacementImage.ps1')
. (Join-Path $PSScriptRoot 'New-NativeReplacementJpeg.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeCharacterAttributeControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeParagraphFormatControl.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeParagraphTopologyProfiles.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeMixedRunTextProfiles.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeFormattingTopologyProfiles.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeEmptyRefillProfiles.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeRunSegmentationProfiles.ps1')
. (Join-Path $PSScriptRoot 'Set-NativeMixedRunFormatting.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeEmptyParagraphTypingStyles.ps1')
. (Join-Path $PSScriptRoot 'Set-NativeParagraphTopologyFormatting.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeParagraphTopologyControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeTableParagraphTopologyControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeFrameLayoutControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeFrameBreakControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeLabelFittingControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeHyphenationControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeVerticalPlacementControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeFontMetricControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeFallbackAlignmentControl.ps1')
. (Join-Path $PSScriptRoot 'Add-NativeTableControl.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeTableGeometry.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeTableCellInventory.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeShapeInventory.ps1')
. (Join-Path $PSScriptRoot 'Get-NativeInventory.ps1')
. (Join-Path $PSScriptRoot 'Write-NativePublisherPhaseReceipt.ps1')
. (Join-Path $PSScriptRoot 'Get-NativePublisherAutomationProcess.ps1')

function Set-NativePhase {
    param([string] $Value)
    $script:phase = $Value
    $script:phaseSequence++
    try {
        if ($nativeVerified) {
            $processReceipt.phase = $Value
            $processReceipt.phaseSequence = $phaseSequence
            Write-NativePublisherPhaseReceipt -Path $ReceiptPath -Value $processReceipt
        } else {
            $startup = [ordered]@{
                schema = 'officeimo.publisher.native-startup.v1'
                workerProcessId = $PID
                controllerProcessId = $ControllerProcessId
                phase = $Value
                phaseSequence = $phaseSequence
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
    try {
        if (-not $nativeVerified -or $nativeId -ne $ownedProcess.Id -or
            $nativeProcess.StartTime.ToUniversalTime().Ticks -ne $ownedProcess.StartTime.ToUniversalTime().Ticks -or
            $nativeProcess.Path -ine $ownedProcess.Path) {
            throw 'The COM window does not match the recorded isolated Publisher automation server.'
        }
    } finally {
        $nativeProcess.Dispose()
    }
    $window.Visible = $false
    return $true
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
    Set-NativePhase 'IdentifyAutomationProcess'
    $ownedProcess = Get-NativePublisherAutomationProcess -StartedAfterUtc $operationStart
    $processReceipt = [ordered]@{
        schema = 'officeimo.publisher.native-process.v1'
        workerProcessId = $PID
        controllerProcessId = $ControllerProcessId
        processId = $ownedProcess.Id
        startTimeUtcTicks = $ownedProcess.StartTime.ToUniversalTime().Ticks.ToString([Globalization.CultureInfo]::InvariantCulture)
        executable = $ownedProcess.Path
        executableVersion = (Get-Item -LiteralPath $ownedProcess.Path).VersionInfo.FileVersion
        executableSha256 = (Get-FileHash -LiteralPath $ownedProcess.Path -Algorithm SHA256).Hash.ToLowerInvariant()
        case = $Case
        phase = 'ApplicationIdentified'
        phaseSequence = $phaseSequence
        identitySource = 'NewIsolatedAutomationServer'
        workerPowerShellVersion = $PSVersionTable.PSVersion.ToString()
        operationSucceeded = $false
    }
    Write-NewNativeJson -Path $ReceiptPath -Value $processReceipt
    $nativeVerified = $true
    Set-NativePhase 'SetAutomationSecurity'
    $application.AutomationSecurity = 3
    if ($Operation -eq 'Corpus') {
        Set-NativePhase 'CreateDocument'
        $document = $application.NewDocument()
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
        Set-NativePhase 'OpenCandidateReadOnly'
        $document = $application.Open($snapshot, $true, $false)
        Set-NativePhase 'ObserveCandidateWindow'
        $null = Register-NativeWindowIdentity -AllowNoWindow
        Set-NativePhase 'InitialInventory'
        $initialInventory = Get-NativeInventory -Publication $document -Objects $objects
        if (-not $initialInventory.readOnly) {
            throw 'Publisher did not open the initial candidate read-only.'
        }
        $qualifiedOperations.Add('ReadOnlyOpen')
        $qualifiedOperations.Add('InitialInventory')
        $producer = [ordered]@{
            name = 'Microsoft Publisher'
            version = [string]$application.Version
            executableVersion = $processReceipt.executableVersion
            executableSha256 = $processReceipt.executableSha256
            powershell = $PSVersionTable.PSVersion.ToString()
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
        if ($Stage -in @('Export', 'Full')) {
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
        $nativeTableClusterMutation = $null
        if ($Stage -in @('Resave', 'Full')) {
            # Only this task-owned snapshot is opened writable. The original
            # candidate remains outside the native application's mutation path.
            Set-NativePhase 'OpenSnapshotWritable'
            $document = $application.Open($snapshot, $false, $false)
            if ($nativeEditRequested) {
                Set-NativePhase 'NativeTextMutation'
                $nativeTextMutation = Set-NativeCandidateText -Publication $document -Objects $objects -Target $NativeEditTarget -Text $NativeEditText
                $qualifiedOperations.Add($(switch ($NativeEditTarget) {
                    'TextBox' { 'NativeTextMutation' }
                    'FirstTableCell' { 'NativeTableCellTextMutation' }
                    'FirstMergedTableCell' { 'NativeMergedTableCellTextMutation' }
                    'LinkedStoryFirstCharacter' { 'NativeLinkedStoryFirstCharacterMutation' }
                }))
            }
            if ($NativeAddClusterTable) {
                Set-NativePhase 'NativeTableClusterAddition'
                $nativeTableClusterMutation = Add-NativeCandidateClusterTable -Publication $document -Objects $objects
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
            $resavedInventory = Get-NativeInventory -Publication $document -Objects $objects
            if (-not $resavedInventory.readOnly) {
                throw 'Publisher did not reopen the native resave read-only.'
            }
            if ($nativeEditRequested) {
                $editedText = $null
                if ($NativeEditTarget -eq 'LinkedStoryFirstCharacter') {
                    $reopenedState = Get-NativeLinkedStoryEditState -Publication $document -Objects $objects -ExpectedText $nativeTextMutation.afterState.text
                    Assert-NativeLinkedStoryEditState -Before $nativeTextMutation.beforeState -After $reopenedState.Snapshot
                    $nativeTextMutation.reopenedState = $reopenedState.Snapshot
                } elseif ($NativeEditTarget -eq 'TextBox') {
                    if ($resavedInventory.pageCount -eq 1 -and $resavedInventory.pages[0].shapes.Count -eq 1) {
                        $editedText = $resavedInventory.pages[0].shapes[0].text
                    }
                } elseif ($NativeEditTarget -in @('FirstTableCell', 'FirstMergedTableCell') -and $resavedInventory.pageCount -eq 1) {
                    $tables = @($resavedInventory.pages[0].shapes | Where-Object type -eq 18)
                    if ($tables.Count -eq 1) {
                        $cells = @($tables[0].tableCells | Where-Object { $_.row -eq 1 -and $_.column -eq 1 })
                        $expectedColumns = if ($NativeEditTarget -eq 'FirstMergedTableCell') { 2 } else { 1 }
                        if ($cells.Count -eq 1 -and $cells[0].rowSpan -eq 1 -and $cells[0].columnSpan -eq $expectedColumns) { $editedText = $cells[0].text }
                    }
                }
                if ($NativeEditTarget -ne 'LinkedStoryFirstCharacter' -and $editedText -cne ($NativeEditText + "`r")) { throw 'The native text mutation did not survive save-as and reopen.' }
            }
            if ($NativeAddClusterTable) {
                if ($resavedInventory.pageCount -ne 1) {
                    throw 'The native cluster-table addition did not retain one physical page.'
                }
                $initialTables = @($initialInventory.pages[0].shapes | Where-Object type -eq 18)
                $resavedTables = @($resavedInventory.pages[0].shapes | Where-Object type -eq 18)
                $retainedTables = @($resavedTables | Where-Object {
                    $_.name -ceq $nativeTableClusterMutation.retainedTableName -and
                    $_.tableGeometry.rowHeightsPoints.Count -eq 18 -and
                    $_.tableGeometry.columnWidthsPoints.Count -eq 18
                })
                $probeTables = @($resavedTables | Where-Object name -CEQ $nativeTableClusterMutation.shapeName)
                if ($initialTables.Count -ne 1 -or $resavedTables.Count -ne 2 -or $retainedTables.Count -ne 1 -or
                    $probeTables.Count -ne 1 -or $probeTables[0].tableGeometry.rowHeightsPoints.Count -ne 2 -or
                    $probeTables[0].tableGeometry.columnWidthsPoints.Count -ne 4 -or $probeTables[0].tableCells.Count -ne 8) {
                    throw 'The native cluster-table addition did not retain its original and named probe tables.'
                }
                for ($row = 1; $row -le 2; $row++) {
                    for ($column = 1; $column -le 4; $column++) {
                        $cells = @($probeTables[0].tableCells | Where-Object { $_.row -eq $row -and $_.column -eq $column })
                        $expectedCellText = $nativeTableClusterMutation.expectedTexts[($row - 1) * 4 + $column - 1] + "`r"
                        if ($cells.Count -ne 1 -or $cells[0].rowSpan -ne 1 -or $cells[0].columnSpan -ne 1 -or
                            $cells[0].text -cne $expectedCellText) {
                            throw ('The native cluster-table addition did not retain probe cell {0},{1}.' -f $row, $column)
                        }
                    }
                }
                $qualifiedOperations.Add('NativeTableClusterAddition')
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
        } elseif ($Stage -eq 'Export') {
            $status = 'OpenedReadOnlyInventoriedAndExported'
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
        if ($NativeAddClusterTable) {
            $report.nativeTableClusterMutation = $nativeTableClusterMutation
        }
        Write-NewNativeJson -Path $validationPath -Value $report
        Set-NativePhase 'Shutdown'
        return
    }
    Set-NativePhase 'AuthorCorpusCase'
    $document.PageSetup.PageWidth = $width
    $document.PageSetup.PageHeight = $height
    $page = $document.Pages.Item(1)
    $objects.Add($page)
    switch ($Case) {
        { $_ -in @('CaptionTypingPreRange', 'CaptionTypingFullRange', 'CaptionTypingExplicit', 'CaptionTypingPreRangeRefilled', 'CaptionTypingFullRangeRefilled', 'CaptionTypingExplicitRefilled') } {
            $mode = if ($Case -like 'CaptionTypingPreRange*') { 'PreRange' } elseif ($Case -like 'CaptionTypingFullRange*') { 'FullRange' } else { 'Explicit' }
            $captionRangeObservations = Add-NativeCaptionTypingControl -FirstPage $page -Mode $mode -Refilled:($Case.EndsWith('Refilled', [StringComparison]::Ordinal))
            $pageCount = 1
        }
        'Blank' { }
        'BlankPages' {
            $pageCount = 2
            $second = $document.Pages.Add(1, 1)
            $objects.Add($second)
        }
        'TextPages' {
            $pageCount = 3
            $middle = $document.Pages.Add(1, 1)
            $objects.Add($middle)
            $last = $document.Pages.Add(1, 2)
            $objects.Add($last)
            $firstShape = Add-TextFrame $page 'FirstPageText' 'First page: native authoring' 72 72 288 108 12
            $lastShape = Add-TextFrame $last 'LastPageText' 'Last page: native authoring' 108 144 360 90 18
            $firstShape.TextFrame.ColumnSpacing = 0
            $lastShape.TextFrame.ColumnSpacing = 0
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
        'TextNonuniformParagraphs' {
            $pageCount = Add-NativeParagraphTopologyControl -Publication $document -FirstPage $page
        }
        'TextRunSegmentation' {
            $pageCount = Add-NativeParagraphTopologyControl -Publication $document -FirstPage $page -RunSegmentation
        }
        'TextEmptyRefill' {
            $pageCount = Add-NativeParagraphTopologyControl -Publication $document -FirstPage $page -EmptyRefill
        }
        'TextFormattingTopology' {
            $pageCount = Add-NativeParagraphTopologyControl -Publication $document -FirstPage $page -FormattingChanges
        }
        'TextMixedRunTextReplacements' {
            $pageCount = Add-NativeParagraphTopologyControl -Publication $document -FirstPage $page -MixedRuns
        }
        'TextNonuniformTextReplacements' {
            $pageCount = Add-NativeParagraphTopologyControl -Publication $document -FirstPage $page -TextReplacements
        }
        'TextNonuniformEmptyParagraphs' {
            $pageCount = Add-NativeParagraphTopologyControl -Publication $document -FirstPage $page -EmptyParagraphs
        }
        'TextUniformEmptyParagraphs' {
            $pageCount = Add-NativeParagraphTopologyControl -Publication $document -FirstPage $page -EmptyParagraphs -UniformEmpty
        }
        'TableNonuniformParagraphs' {
            $pageCount = Add-NativeTableParagraphTopologyControl -Publication $document -FirstPage $page
        }
        'TableRunSegmentation' {
            $pageCount = Add-NativeTableParagraphTopologyControl -Publication $document -FirstPage $page -RunSegmentation
        }
        'TableEmptyRefill' {
            $pageCount = Add-NativeTableParagraphTopologyControl -Publication $document -FirstPage $page -EmptyRefill
        }
        'TableFormattingTopology' {
            $pageCount = Add-NativeTableParagraphTopologyControl -Publication $document -FirstPage $page -FormattingChanges
        }
        'TableMixedRunTextReplacements' {
            $pageCount = Add-NativeTableParagraphTopologyControl -Publication $document -FirstPage $page -MixedRuns
        }
        'TableNonuniformTextReplacements' {
            $pageCount = Add-NativeTableParagraphTopologyControl -Publication $document -FirstPage $page -TextReplacements
        }
        'TableNonuniformEmptyParagraphs' {
            $pageCount = Add-NativeTableParagraphTopologyControl -Publication $document -FirstPage $page -EmptyParagraphs
        }
        'TableUniformEmptyParagraphs' {
            $pageCount = Add-NativeTableParagraphTopologyControl -Publication $document -FirstPage $page -EmptyParagraphs -UniformEmpty
        }
        'TextParagraphAttributes' {
            $pageCount = Add-NativeParagraphFormatControl -Publication $document -FirstPage $page
        }
        'TextVerticalPlacement' {
            $pageCount = Add-NativeVerticalPlacementControl -Publication $document -FirstPage $page
        }
        'TextFontMetrics' {
            $pageCount = Add-NativeFontMetricControl -Publication $document -FirstPage $page
        }
        'TextFallbackAlignment' {
            $pageCount = Add-NativeFallbackAlignmentControl -Publication $document -FirstPage $page
        }
        'TextFrameBreaks' {
            $pageCount = Add-NativeFrameBreakControl -Publication $document -FirstPage $page
        }
        'TextLabelFitting' {
            $pageCount = Add-NativeLabelFittingControl -Publication $document -FirstPage $page
        }
        'TextHyphenation' {
            $pageCount = Add-NativeHyphenationControl -Publication $document -FirstPage $page
        }
        'TextFrameLayout' {
            $pageCount = Add-NativeFrameLayoutControl -Publication $document -FirstPage $page
        }
        'TextAlignedTabs' {
            $pageCount = Add-NativeParagraphFormatControl -Publication $document -FirstPage $page -AlignedTabs
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
        { $_ -in @('ObjectPages', 'ObjectPagesEdited') } {
            $picturePath = Join-Path $output ($stem + '.png')
            $pageCount = Add-NativeObjectPagesControl -Document $document -FirstPage $page -ImagePath $picturePath -Edited:($Case -eq 'ObjectPagesEdited')
        }
        'Image' {
            $picturePath = Join-Path $output ($stem + '.png')
            New-ReferencePicture $picturePath
            $image = $page.Shapes.AddPicture($picturePath, 0, -1, 72, 72, 96, 96)
            $image.Name = 'MinimalImage'
            $objects.Add($image)
            $expectedShapes.Add([pscustomobject]@{ name = 'MinimalImage'; kind = 'Image'; left = 72; top = 72; width = 96; height = 96 })
        }
        { $_ -in @('LinkedStory', 'LinkedStoryOverflow') } {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $pageCount = Add-NativeLinkedStoryControl -Document $document -FirstPage $page -Overflow:($Case -eq 'LinkedStoryOverflow')
        }
        { $_ -in @('MultiParagraphCreation', 'MultiParagraphCreationLinked') } {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $multiParagraphObservations = Add-NativeMultiParagraphControl -Document $document -FirstPage $page -Linked:($Case -eq 'MultiParagraphCreationLinked')
            $pageCount = $multiParagraphObservations.pageCount
        }
        { $_ -in @('FrameCreation', 'FrameCreationEdited') } {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $pageCount = Add-NativeFrameCreationControl -Document $document -FirstPage $page -Edited:($Case -eq 'FrameCreationEdited')
        }
        'FrameCreationFrench' {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $pageCount = Add-NativeFrameCreationControl -Document $document -FirstPage $page -LanguageId 1036
        }
        { $_ -in @('TableClusterNear', 'TableClusterNearEdited', 'TableClusterWide') } {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $size = if ($Case -eq 'TableClusterWide') { 20 } else { 18 }
            $pageCount = Add-NativeTableClusterControl -FirstPage $page -Size $size -Edited:($Case -eq 'TableClusterNearEdited')
        }
        { $_ -in @('TableUtilityPair', 'TableUtilityPairStyled', 'TableUtilityPairBorders', 'TableUtilityPairFills', 'TableUtilityPairComposed', 'TableUtilityPairSingleFill', 'TableUtilityPairTargetComposed') } {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $kind = switch ($Case) {
                'TableUtilityPairStyled' { 'Styled' }
                'TableUtilityPairBorders' { 'Borders' }
                'TableUtilityPairFills' { 'Fills' }
                'TableUtilityPairComposed' { 'Composed' }
                'TableUtilityPairSingleFill' { 'SingleFill' }
                'TableUtilityPairTargetComposed' { 'TargetComposed' }
                default { 'Bare' }
            }
            $pageCount = Add-NativeTableUtilityPairControl -FirstPage $page -Kind $kind
        }
        { $_ -in @('TableDecoration', 'TableDecorationEdited', 'TableDecorationText', 'TableBordersOnly', 'TableBordersOnlyText', 'TableFillsOnly', 'TableBordersOnlyCleared', 'TableFillVisibility', 'TableFillVisibilityRestored', 'TableFillVisibilityText', 'TableBorderTopology', 'TableBordersCleared', 'TableBorderTopologyTextFill', 'TableBorderDefaultFill') } {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $pageCount = Add-NativeTableDecorationControl -FirstPage $page -Edited:($Case -eq 'TableDecorationEdited') -TextEdited:($Case -in @('TableDecorationText', 'TableFillVisibilityText', 'TableBorderTopologyTextFill', 'TableBordersOnlyText')) -FillVisibility:($Case -in @('TableFillVisibility', 'TableFillVisibilityText', 'TableBorderTopologyTextFill')) -RestoreFillVisibility:($Case -eq 'TableFillVisibilityRestored') -BorderTopology:($Case -in @('TableBorderTopology', 'TableBorderTopologyTextFill', 'TableBorderDefaultFill')) -ClearBorders:($Case -in @('TableBordersCleared', 'TableBordersOnlyCleared')) -DefaultCellFill:($Case -eq 'TableBorderDefaultFill') -BorderOnly:($Case -in @('TableBordersOnly', 'TableBordersOnlyText', 'TableBordersOnlyCleared')) -FillOnly:($Case -eq 'TableFillsOnly')
        }
        { $_ -in @('TableMergedDecoration', 'TableMergedDecorationEdited', 'TableMergedDecorationText', 'TableMergedDecorationComposed', 'TableMergedFills', 'TableMergedBorders') } {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $variant = switch ($Case) {
                'TableMergedDecorationEdited' { 'Edited' }
                'TableMergedDecorationText' { 'Text' }
                'TableMergedDecorationComposed' { 'Composed' }
                'TableMergedFills' { 'Fills' }
                'TableMergedBorders' { 'Borders' }
                default { 'Base' }
            }
            $pageCount = Add-NativeMergedTableDecorationControl -FirstPage $page -Variant $variant
        }
        { $_ -in @('TableCreation', 'TableCreationEdited', 'TableCreationMerged', 'TableCreationMergedEdited') } {
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $pageCount = Add-NativeTableCreationControl -FirstPage $page -Edited:($Case -in @('TableCreationEdited', 'TableCreationMergedEdited')) -Merged:($Case -in @('TableCreationMerged', 'TableCreationMergedEdited'))
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
        { $_ -in @('TableLabel', 'TableLabelWrapped', 'TableLabelShortParagraphs', 'TableCellTyping', 'TableCellTypingEarlyRange', 'TableCellParagraphs') } {
            $pageCount = Add-NativeTableControl -Publication $document -FirstPage $page -Case $Case -PicturePath (Join-Path $output ($stem + '.png'))
        }
        { $_ -in @('Brochure', 'MultipleMasters', 'TwoPageMaster', 'MultipleMastersEdited', 'TwoPageMasterEdited', 'MultipleMastersReassigned', 'MultipleMastersIgnored', 'TwoPageMasterReassigned', 'TwoPageMasterIgnored') } {
            $layout = Add-NativeComparisonLayout -Case $Case -Document $document -FirstPage $page
            $pageCount = $layout.pages
        }
        { $_ -in @('MasterContent', 'MasterContentEdited') } {
            $picturePath = Join-Path $output ($stem + '.png')
            $replacementPath = Join-Path $output ($stem + '-replacement.png')
            $pageCount = Add-NativeMasterContentControl -Document $document -FirstPage $page -ImagePath $picturePath -ReplacementPath $replacementPath -Edited:($Case -eq 'MasterContentEdited')
        }
        { $_ -in @('WrapObjectOrderDefaults', 'WrapObjectOrderLayout', 'WrapTextFrameOutside', 'WrapObjectOrderLeft', 'WrapObjectOrderRight', 'WrapObjectOrderLarger', 'WrapObjectOrderLargerLeft', 'WrapObjectOrderLargerTie', 'WrapObjectOrderLargerColumns', 'WrapObjectOrderOutside', 'WrapObjectOrderRectangleOnly', 'WrapObjectOrderPictureOnly', 'WrapObjectOrderSwap', 'WrapObjectOrder', 'WrapObjectOrderReversed', 'WrapObjectOrderNone') } {
            $picturePath = Join-Path $output ($stem + '.png')
            $order = switch ($Case) {
                'WrapObjectOrderReversed' { 'Reverse' }
                'WrapObjectOrderRectangleOnly' { 'RectangleOnly' }
                'WrapObjectOrderPictureOnly' { 'PictureOnly' }
                'WrapObjectOrderSwap' { 'Swap' }
                default { 'Original' }
            }
            $wrapSide = switch ($Case) { 'WrapObjectOrderLeft' { 1 }; 'WrapObjectOrderRight' { 2 }; { $_ -in @('WrapObjectOrderLarger', 'WrapObjectOrderLargerLeft', 'WrapObjectOrderLargerTie', 'WrapObjectOrderLargerColumns') } { 3 }; default { 0 } }
            $largerPlacement = switch ($Case) { 'WrapObjectOrderLargerLeft' { 'Left' }; 'WrapObjectOrderLargerTie' { 'Tie' }; default { 'Default' } }
            $pageCount = Add-NativeWrapOrderControl -FirstPage $page -ImagePath $picturePath -Order $order -NoWrap:($Case -eq 'WrapObjectOrderNone') -Side $wrapSide -LargerPlacement $largerPlacement -Columns:($Case -eq 'WrapObjectOrderLargerColumns') -Outside:($Case -eq 'WrapObjectOrderOutside') -FrameOutside:($Case -eq 'WrapTextFrameOutside') -Layout:($Case -eq 'WrapObjectOrderLayout') -UseDefaults:($Case -eq 'WrapObjectOrderDefaults')
        }
        { $_ -in @('MixedObjectPages', 'MixedObjectPagesEdited', 'MixedObjectOrder', 'MixedObjectOrderReversed') } {
            $picturePath = Join-Path $output ($stem + '.png')
            $pageCount = Add-NativeMixedPagesControl -Document $document -FirstPage $page -ImagePath $picturePath -Edited:($Case -eq 'MixedObjectPagesEdited') -ReverseOrder:($Case -eq 'MixedObjectOrderReversed')
        }
        { $_ -in @('ImagePages', 'ImagePagesEdited', 'ImagePagesReplacement', 'ImagePagesJpegReplacement', 'ImagePagesJpegCropped', 'ImagePagesCropped', 'SharedImagePages', 'SharedImagePagesReplacement') } {
            $picturePath = Join-Path $output ($stem + '.png')
            $replacementPath = Join-Path $output ($stem + '-replacement.png')
            $jpegPath = if ($Case -in @('ImagePagesJpegReplacement', 'ImagePagesJpegCropped')) { Join-Path $output ($stem + '-replacement.jpg') } else { $null }
            $pageCount = Add-NativeImagePagesControl -Document $document -FirstPage $page -ImagePath $picturePath -SecondImagePath $replacementPath -Edited:($Case -eq 'ImagePagesEdited') -ReplaceLast:($Case -eq 'ImagePagesReplacement') -ReplacementJpegPath $jpegPath -Cropped:($Case -eq 'ImagePagesCropped') -CropLast:($Case -eq 'ImagePagesJpegCropped') -SharedImages:($Case -in @('SharedImagePages', 'SharedImagePagesReplacement')) -ReplaceSharedFirst:($Case -eq 'SharedImagePagesReplacement')
        }
        'PictureOwnership' {
            $picturePath = Join-Path $output ($stem + '.png')
            New-ReferencePicture $picturePath
            $layout = Add-NativeComparisonPictures -Document $document -FirstPage $page -PicturePath $picturePath
            $pageCount = $layout.pages
        }
        'PictureReplacement' {
            $picturePath = Join-Path $output ($stem + '.png')
            $replacementPath = Join-Path $output ($stem + '-replacement.png')
            New-ReferencePicture $picturePath
            New-NativeReplacementImage $replacementPath
            $layout = Add-NativeComparisonPictures -Document $document -FirstPage $page -PicturePath $picturePath -ReplacementPath $replacementPath
            $pageCount = $layout.pages
        }
        'PictureCropFractionsReplacement' {
            $picturePath = Join-Path $output ($stem + '.png')
            $replacementPath = Join-Path $output ($stem + '-replacement.png')
            New-ReferencePicture $picturePath
            New-NativeReplacementImage $replacementPath
            $layout = Add-NativeComparisonPictures -Document $document -FirstPage $page -PicturePath $picturePath -ReplacementPath $replacementPath -FractionalCrop -PhysicalBounds @(84, 132, 120, 84)
            $pageCount = $layout.pages
        }
        'PictureCropFractions' {
            $picturePath = Join-Path $output ($stem + '.png')
            New-ReferencePicture $picturePath
            $layout = Add-NativeComparisonPictures -Document $document -FirstPage $page -PicturePath $picturePath -FractionalCrop
            $pageCount = $layout.pages
        }
        'PictureCropped' {
            $picturePath = Join-Path $output ($stem + '.png')
            New-ReferencePicture $picturePath
            $layout = Add-NativeComparisonPictures -Document $document -FirstPage $page -PicturePath $picturePath -Cropped
            $pageCount = $layout.pages
        }
    }
    Set-NativePhase 'ObserveCorpusPageCount'
    if ($document.Pages.Count -ne $pageCount) {
        throw ('Native page count differs from the declared contract: ' + $document.Pages.Count)
    }
    Set-NativePhase 'InventoryCorpusCase'
    $inventory = Get-NativeInventory -Publication $document -Objects $objects
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
    if ($null -ne $captionRangeObservations) {
        Set-NativePhase 'ObserveCaptionAfterSave'
        $captionShapes = $page.Shapes
        $objects.Add($captionShapes)
        $captionShape = $captionShapes.Item('CaptionTyping')
        $objects.Add($captionShape)
        $captionRangeObservations.observations.Add((Get-NativeCaptionRangeObservation -Shape $captionShape -Stage 'AfterSave'))
    }
    if ($null -ne $multiParagraphObservations) {
        Set-NativePhase 'ObserveMultiParagraphAfterSave'
        $multiParagraphShapes = $page.Shapes
        $objects.Add($multiParagraphShapes)
        $multiParagraphShape = $multiParagraphShapes.Item($multiParagraphObservations.firstFrameName)
        $objects.Add($multiParagraphShape)
        $multiParagraphObservations.observations.Add((Get-NativeMultiParagraphObservation -Shape $multiParagraphShape -ExpectedText $multiParagraphObservations.expectedText -Stage 'AfterSave'))
    }
    Set-NativePhase 'ExportCorpusReference'
    $document.ExportAsFixedFormat(2, $reference)
    Set-NativePhase 'ObserveCorpusProducer'
    $producerVersion = [string]$application.Version
    $executableVersion = $processReceipt.executableVersion
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
            executableSha256 = $processReceipt.executableSha256
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
    if ($null -ne $captionRangeObservations) {
        $contract.captionRangeObservations = $captionRangeObservations
    }
    if ($null -ne $multiParagraphObservations) {
        $contract.multiParagraphObservations = $multiParagraphObservations
    }
    if ($null -ne $layout) {
        $contract.expected.masterAssignments = $layout.masterAssignments
        $contract.expected.panelWidthPoints = $layout.panelWidthPoints
    }
    if ($Case -in @('PictureOwnership', 'PictureReplacement', 'PictureCropped', 'PictureCropFractions', 'PictureCropFractionsReplacement')) {
        $contract.nativePictureObservations = $layout.pictureObservations
        if ($Case -in @('PictureCropFractions', 'PictureCropFractionsReplacement')) { $contract.expected.sourceCropBasis = $layout.sourceCropBasis }
        $contract.expected.pictureDuplication = [pscustomobject]@{
            sourceName = 'PhysicalPictureOriginal'
            duplicateName = 'PhysicalPictureDuplicate'
            method = 'Shape.Duplicate'
            nativeImageStoreSharing = 'Unassessed'
        }
    }
    if ($Case -in @('ImagePages', 'ImagePagesEdited', 'ImagePagesReplacement', 'ImagePagesJpegReplacement', 'ImagePagesJpegCropped', 'ImagePagesCropped', 'SharedImagePages', 'SharedImagePagesReplacement')) {
        $contract.files.secondSourceImage = [ordered]@{
            name = [IO.Path]::GetFileName($replacementPath)
            sha256 = (Get-FileHash -LiteralPath $replacementPath -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    if ($Case -in @('PictureReplacement', 'PictureCropFractionsReplacement', 'ImagePagesReplacement', 'ImagePagesJpegReplacement', 'ImagePagesJpegCropped')) {
        $replacementAsset = if ($Case -in @('PictureReplacement', 'PictureCropFractionsReplacement')) { $replacementPath }
            elseif ($Case -eq 'ImagePagesReplacement') { $picturePath }
            else { $jpegPath }
        $contract.expected.pictureReplacement = [pscustomobject]@{
            targetName = if ($Case -in @('PictureReplacement', 'PictureCropFractionsReplacement')) { 'PhysicalPictureOriginal' } else { 'LastPageImage' }
            method = 'PictureFormat.Replace'
            insertAs = 'Embedded'
            framePolicy = if ($Case -in @('PictureCropFractionsReplacement', 'ImagePagesJpegCropped')) {
                'Restore authored bounds and apply the declared source crop after replacement.'
            } else { 'Restore authored bounds and zero crop after replacement.' }
        }
        $contract.files.replacementImage = [ordered]@{
            name = [IO.Path]::GetFileName($replacementAsset)
            sha256 = (Get-FileHash -LiteralPath $replacementAsset -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    if ($Case -in @('MasterContent', 'MasterContentEdited')) {
        $contract.expected.wrapType = 0
        $contract.expected.masterObjects = @('MasterBackground', 'MasterImage', 'MasterText')
        if ($Case -eq 'MasterContentEdited') {
            $contract.files.replacementImage = [ordered]@{ name = [IO.Path]::GetFileName($replacementPath); sha256 = (Get-FileHash -LiteralPath $replacementPath -Algorithm SHA256).Hash.ToLowerInvariant() }
        }
    }
    if ($Case -in @('WrapObjectOrderDefaults', 'WrapObjectOrderLayout', 'WrapTextFrameOutside', 'WrapObjectOrderLeft', 'WrapObjectOrderRight', 'WrapObjectOrderLarger', 'WrapObjectOrderLargerLeft', 'WrapObjectOrderLargerTie', 'WrapObjectOrderLargerColumns', 'WrapObjectOrderOutside', 'WrapObjectOrderRectangleOnly', 'WrapObjectOrderPictureOnly', 'WrapObjectOrderSwap', 'WrapObjectOrder', 'WrapObjectOrderReversed', 'WrapObjectOrderNone')) {
        $contract.expected.wrap = [ordered]@{ type = if ($Case -eq 'WrapObjectOrderNone') { 0 } else { 1 }; side = if ($Case -ne 'WrapObjectOrderNone') { switch ($Case) { 'WrapObjectOrderLeft' { 1 }; 'WrapObjectOrderRight' { 2 }; { $_ -in @('WrapObjectOrderLarger', 'WrapObjectOrderLargerLeft', 'WrapObjectOrderLargerTie', 'WrapObjectOrderLargerColumns') } { 3 }; default { 0 } } }; automaticDistances = if ($Case -ne 'WrapObjectOrderNone') { $false }; distancePoints = if ($Case -ne 'WrapObjectOrderNone') { 6 }; textFrameWrapType = 0 }
        if ($Case -eq 'WrapObjectOrderDefaults') { $contract.expected.wrap = [ordered]@{ obstacleSettings = 'Untouched defaults; observed types recorded per shape'; textFrameWrapType = 0 } }
        $contract.expected.drawingOrder = switch ($Case) {
            'WrapObjectOrderReversed' { @('WrapPicture', 'WrapRectangle', 'WrappedStory') }
            'WrapObjectOrderRectangleOnly' { @('WrapPicture', 'WrappedStory', 'WrapRectangle') }
            'WrapObjectOrderPictureOnly' { @('WrapRectangle', 'WrappedStory', 'WrapPicture') }
            'WrapObjectOrderSwap' { @('WrappedStory', 'WrapPicture', 'WrapRectangle') }
            default { @('WrappedStory', 'WrapRectangle', 'WrapPicture') }
        }
    }
    if ($Case -in @('MixedObjectPages', 'MixedObjectPagesEdited', 'MixedObjectOrder', 'MixedObjectOrderReversed')) {
        $contract.expected.wrapType = 0
        $contract.expected.drawingOrder = @(@('FirstBackground', 'FirstImage', 'FirstText', 'FirstBand'), @(), @('LastText', 'LastImage', 'LastBand', 'TopImage'))
        if ($Case -eq 'MixedObjectOrderReversed') {
            $contract.expected.drawingOrder = @(@('FirstBand', 'FirstText', 'FirstImage', 'FirstBackground'), @(), @('TopImage', 'LastBand', 'LastImage', 'LastText'))
        }
    }
    if ($Case -in @('SharedImagePages', 'SharedImagePagesReplacement')) { $contract.expected.sharedImagePlacements = @('FirstPageImage', 'SharedLastPageImage') }
    if ($Case -eq 'SharedImagePagesReplacement') { $contract.expected.firstPictureReplacement = 'secondSourceImage'; $contract.expected.sharedImagePlacements = @('SharedLastPageImage') }
    if ($Case -eq 'ImagePagesCropped') { $contract.expected.firstPictureCropFractions = @(0.25, 0.125, 0.0625, 0.1875) }
    if ($Case -eq 'ImagePagesJpegCropped') { $contract.expected.lastPictureCropFractions = @(0.25, 0.125, 0.0625, 0.1875) }
    $sourceImage = Join-Path $output ($stem + '.png')
    if (Test-Path -LiteralPath $sourceImage) {
        $contract.files.sourceImage = [ordered]@{ name = [IO.Path]::GetFileName($sourceImage); sha256 = (Get-FileHash -LiteralPath $sourceImage -Algorithm SHA256).Hash.ToLowerInvariant() }
    }
    Write-NewNativeJson -Path $contractPath -Value $contract
    Set-NativePhase 'Shutdown'
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
    if ($nativeVerified) {
        try {
            $currentBinaryHash = (Get-FileHash -LiteralPath $processReceipt.executable -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($currentBinaryHash -ne $processReceipt.executableSha256) {
                throw 'The Publisher executable changed during qualification; its runtime attribution is unassessed.'
            }
        } catch {
            $cleanupFailures.Add('Executable identity: ' + $_.Exception.Message)
        }
    }
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
