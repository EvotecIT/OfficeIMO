<#
.SYNOPSIS
Creates independent Publisher publications and native PDF references for comparison.
.DESCRIPTION
Runs each selected case in an isolated STA PowerShell process with a bounded deadline.
Microsoft Publisher must be installed on Windows and no Publisher process may already
be running. The tool creates synthetic publications from scratch; it does not open
user documents, enable macros, change Office settings, or invoke OfficeIMO.
The output directory must be empty. Native output is evidence, not an OfficeIMO
writer result or a claim of matching layout. Native authoring is qualified with
PowerShell 7. PowerShell 5.1 syntax and receipt checks do not qualify native COM saves.
.PARAMETER OutputDirectory
Empty task-owned directory for publications, PDF references, contracts and run receipts.
.PARAMETER Case
Cases to produce. Text and TextLong differ only in their text length; Blank, Rectangle
and Image isolate the smallest native structures needed by a template-free writer.
TextParagraphs isolates a single story with two text paragraphs and an empty one.
TextStories separates three stories; TextFormattingPages retains an overflowing
96-paragraph story with multiple native paragraph-formatting pages.
TextFormattingMany has 240 paragraphs; TextCharacterPages uses four independent
font sizes to produce multiple character-formatting pages.
TextCharacterAttributes isolates native fonts, colors, italic, underline,
strikethrough and baseline attributes in twelve independent frames on two pages.
TextUnderlinePatterns records all eighteen automation underline choices on three pages.
TextFrameLayout covers nine column, gap, inset and vertical-alignment profiles on three pages.
ZeroAnchors places each of the four client-anchor coordinates at native zero.
Brochure has two landscape sides with three panels each. MultipleMasters and
TwoPageMaster exercise explicit native single-page and paired master assignments.
PictureOwnership duplicates an embedded picture and applies one master picture
to two physical pages. Native image-store sharing is a separate saved-file check.
PictureReplacement replaces one duplicate with a wide embedded PNG, retaining
the authored frame bounds and zero crop.
PictureCropped retains explicit nonzero crop on one physical placement of the shared image.
TableCellTyping records clearing, styled empty CRs, language and refill on three pages.
TableCellTypingEarlyRange isolates ranges captured before text assignment on two pages.
.PARAMETER TimeoutSeconds
Maximum elapsed time per case, including application startup, save and shutdown.
.EXAMPLE
./Build/Publisher/New-NativeComparisonCorpus.ps1 -OutputDirectory D:/Scratch/publisher-controls
.EXAMPLE
./Build/Publisher/New-NativeComparisonCorpus.ps1 -OutputDirectory D:/Scratch/publisher-minimal -Case Blank,Text,TextLong
.OUTPUTS
One summary object per case. corpus-manifest.json records the selected case contracts.
.NOTES
Only the worker and a Publisher process identified by its COM window and exact
start time are eligible for forced cleanup. A startup hang before that identity
is recorded leaves Publisher ownership uncertain and reports the gap.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputDirectory,

    [ValidateSet('Blank', 'BlankPages', 'Text', 'TextLong', 'TextBlank', 'TextParagraphs', 'TextParagraphAttributes', 'TextAlignedTabs', 'TextFrameLayout', 'TextStories', 'TextFormattingPages', 'TextFormattingMany', 'TextCharacterPages', 'TextCharacterAttributes', 'TextUnderlinePatterns', 'ZeroAnchors', 'Rectangle', 'Image', 'Newsletter', 'TableLabel', 'TableCellTyping', 'TableCellTypingEarlyRange', 'Brochure', 'MultipleMasters', 'TwoPageMaster', 'PictureOwnership', 'PictureReplacement', 'PictureCropped')]
    [string[]] $Case = @('Blank', 'BlankPages', 'Text', 'TextLong', 'TextBlank', 'TextParagraphs', 'TextParagraphAttributes', 'TextAlignedTabs', 'TextFrameLayout', 'TextStories', 'TextFormattingPages', 'TextFormattingMany', 'TextCharacterPages', 'TextCharacterAttributes', 'TextUnderlinePatterns', 'ZeroAnchors', 'Rectangle', 'Image', 'Newsletter', 'TableLabel', 'TableCellTyping', 'TableCellTypingEarlyRange', 'Brochure', 'MultipleMasters', 'TwoPageMaster', 'PictureOwnership', 'PictureReplacement', 'PictureCropped'),

    [ValidateRange(10, 600)]
    [int] $TimeoutSeconds = 90
)

$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') {
    throw 'The independent producer requires Microsoft Publisher on Windows.'
}
if (@(Get-Process -Name MSPUB -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'Publisher is already running. Preserve that session and run this tool after it is closed.'
}
if (@($Case | Select-Object -Unique).Count -ne $Case.Count) {
    throw 'Select each corpus case at most once.'
}
$output = [IO.Path]::GetFullPath($OutputDirectory)
if ((Test-Path -LiteralPath $output) -and @(Get-ChildItem -LiteralPath $output -Force).Count -gt 0) {
    throw 'Choose an empty task-owned output directory; existing evidence is not overwritten.'
}
[void][IO.Directory]::CreateDirectory($output)
$reservationPath = Join-Path $output '.corpus-reservation.json'
$reservationBytes = [Text.UTF8Encoding]::new($false).GetBytes((@{ schema = 'officeimo.publisher.native-corpus-reservation.v1'; controllerProcessId = $PID; cases = $Case } | ConvertTo-Json -Depth 3))
$reservationStream = [IO.FileStream]::new($reservationPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
try {
    $reservationStream.Write($reservationBytes, 0, $reservationBytes.Length)
} finally {
    $reservationStream.Dispose()
}
$controllerPath = Join-Path $PSScriptRoot 'Invoke-NativePublisherOperation.ps1'
$results = [Collections.Generic.List[object]]::new()

foreach ($selected in $Case) {
    $contract = & $controllerPath -Operation Corpus -Case $selected -OutputDirectory $output -TimeoutSeconds $TimeoutSeconds
    if ($contract.case -ne $selected -or $contract.status -ne 'Produced') {
        throw ('The completed case {0} has no valid producer contract.' -f $selected)
    }
    $results.Add($contract)
    [pscustomobject]@{
        Case = $selected
        Status = $contract.status
        Publication = $contract.files.publication.name
        Reference = $contract.files.reference.name
        Pages = $contract.expected.pages
        Contract = Join-Path $output ('native-' + $selected.ToLowerInvariant() + '.json')
    }
}

$manifest = [ordered]@{
    schema = 'officeimo.publisher.native-comparison.v1'
    producer = 'Microsoft Publisher'
    license = 'MIT'
    purpose = 'Independent native inputs and PDF reference exports; not OfficeIMO writer output.'
    cases = $results.ToArray()
}
$json = $manifest | ConvertTo-Json -Depth 14
$manifestBytes = [Text.UTF8Encoding]::new($false).GetBytes($json)
$manifestStream = [IO.FileStream]::new((Join-Path $output 'corpus-manifest.json'), [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
try {
    $manifestStream.Write($manifestBytes, 0, $manifestBytes.Length)
} finally {
    $manifestStream.Dispose()
}
