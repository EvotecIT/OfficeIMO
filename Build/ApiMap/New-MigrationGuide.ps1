#Requires -Version 7.0
<#
.SYNOPSIS
Generates the 3.6 section of MIGRATION.md from the API map.

.DESCRIPTION
Reads powershell/*.json and csharp/*.json next to this script and writes the OfficeIMO and
PSWriteOffice 3.6 upgrade section as Markdown: the global rules, before/after for common jobs,
.NET package changes, and complete reference tables. The output is deterministic, so running it
again on an unchanged map produces an identical file.

The script fails when a map entry does not appear exactly once in the reference tables or when a
generated table row has the wrong number of cells. Map data that looks inconsistent with the API
design (missing hints or reasons, unknown targets, unresolved alternatives) is reported as
warnings; -FailOnDataWarning turns those into an error.

Unknown fields in the map (for example command examples or applied notes) are ignored, except
command "examples", which are shown under the common jobs when present.

.EXAMPLE
./Build/ApiMap/New-MigrationGuide.ps1 -OutputPath ./artefacts/MIGRATION-3.6.md

.EXAMPLE
./Build/ApiMap/New-MigrationGuide.ps1 -OutputPath ./out.md -HeadingLevel 1 -FailOnDataWarning
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputPath,

    [string] $MapDirectory = $PSScriptRoot,

    # Level of the section heading. 2 matches the other sections of MIGRATION.md.
    [ValidateRange(1, 3)]
    [int] $HeadingLevel = 2,

    [switch] $FailOnDataWarning
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

#region Design constants (decisions in Docs/officeimo.api-design.md)

$PowerShellAreaOrder = 'word', 'excel', 'powerpoint', 'pdf', 'markup', 'crossformat', 'other'
$PowerShellAreaTitles = @{
    word        = 'Word'
    excel       = 'Excel'
    powerpoint  = 'PowerPoint'
    pdf         = 'PDF'
    markup      = 'Markdown and markup formats'
    crossformat = 'Cross-format commands'
    other       = 'Other formats'
}
$CSharpPackageOrder = @(
    'OfficeIMO', 'OfficeIMO.Word', 'OfficeIMO.Word.Pdf', 'OfficeIMO.Word.Html', 'OfficeIMO.Word.Markdown',
    'OfficeIMO.Excel', 'OfficeIMO.Excel.Pdf', 'OfficeIMO.PowerPoint', 'OfficeIMO.PowerPoint.Pdf',
    'OfficeIMO.Pdf', 'OfficeIMO.Markdown', 'OfficeIMO.Markdown.Html', 'OfficeIMO.Reader.All', 'OfficeIMO.Workflows'
)
$CSharpCategoryOrder = 'package', 'lifecycle', 'conversion', 'authoring', 'units', 'naming', 'options', 'visibility'
$CSharpCategoryTitles = @{
    package    = 'Package'
    lifecycle  = 'Lifecycle'
    conversion = 'Conversion'
    authoring  = 'Authoring'
    units      = 'Units'
    naming     = 'Naming'
    options    = 'Options'
    visibility = 'Visibility'
}
$PowerShellActionLabels = @{
    'keep'         = 'Kept'
    'rename'       = 'Renamed'
    'merge'        = 'Merged'
    'parameter'    = 'Now a parameter'
    'cross-format' = 'Cross-format'
    'remove'       = 'Removed'
}
$CSharpActionLabels = @{
    'rename'      = 'Renamed'
    'remove'      = 'Removed'
    'move'        = 'Moved'
    'merge'       = 'Merged'
    'change'      = 'Changed'
    'internalize' = 'Internal'
}
# Format names from the command-name rules, with their display spelling.
$FormatNames = [ordered]@{
    OpenDocument = 'OpenDocument'; PowerPoint = 'PowerPoint'; Confluence = 'Confluence'; AsciiDoc = 'AsciiDoc'
    Markdown = 'Markdown'; OneNote = 'OneNote'; Project = 'Project'; Invoice = 'invoices'; Access = 'Access'
    Excel = 'Excel'; Visio = 'Visio'; Email = 'email'; Latex = 'LaTeX'; IWork = 'iWork'; Word = 'Word'
    Html = 'HTML'; Epub = 'EPUB'; Pdf = 'PDF'; Rtf = 'RTF'; Csv = 'CSV'; Ocr = 'OCR'
}
$ConversionFormatNames = @(
    'Word', 'Excel', 'PowerPoint', 'Pdf', 'Markdown', 'Visio', 'Html', 'Rtf', 'OpenDocument', 'Csv', 'Email',
    'Project', 'Access', 'Epub', 'OneNote', 'IWork', 'AsciiDoc', 'Latex', 'Confluence', 'Image', 'Text', 'Xps'
)
$SlideLayouts = 'Title', 'Section', 'TitleAndContent', 'TwoContent', 'Comparison', 'Kpi', 'Cards', 'Process', 'Quote', 'ImageFull', 'Closing', 'CaseStudy', 'Logos', 'Coverage'
$FrontDoorFrameworks = @('net10.0')
$MaxJobExamples = 2

# Common jobs, most useful first. Old names that are missing from the map are skipped with a warning.
$CommonJobs = @(
    [pscustomobject]@{
        Title = 'Export objects to Excel'
        Intro = '`Export-OfficeExcel` keeps its name and becomes styled by default: a themed table with a frozen header and fitted columns. `-NoFreeze`, `-NoAutoFit` and `-NoTableStyle` turn each default off. Typed column switches merge into `-ColumnFormat`, flattening settings into `-FlattenOptions`, document properties into `-DocumentProperty` and save switches into `-SaveOptions`.'
        Old   = @('Export-OfficeExcel')
    }
    [pscustomobject]@{
        Title = 'Read rows from a workbook'
        Intro = '`Import-OfficeExcel` is the one command that reads rows. `-Sheet` takes a name or a 1-based position, and `-Range` and `-Table` narrow the read.'
        Old   = @('Get-OfficeExcelData', 'Get-OfficeExcelRange', 'Import-OfficeExcel')
    }
    [pscustomobject]@{
        Title = 'Open, save and close documents'
        Intro = 'Commands that opened a file with `Get-` are now `Open-`. With a content block, `New-` and `Open-` save and close the file when the block ends. Without a block they return the live document, which you save with `Save-` and close with `Close-`. `-NoSave` is gone, and `-Open` becomes `-Show`.'
        Old   = @('Get-OfficeWord', 'Get-OfficeExcel', 'Get-OfficePowerPoint', 'Save-OfficeExcel')
    }
    [pscustomobject]@{
        Title = 'Write Word documents'
        Intro = 'Styled runs are child commands. Replace `-Run @{ Text = ...; Bold = ... }` hashtables of parallel arrays with `WordText` commands inside `WordParagraph`. The same pattern applies to `PptText`, `PdfText` and `ExcelText`.'
        Old   = @('New-OfficeWord', 'Add-OfficeWordParagraph', 'Add-OfficeWordText')
    }
    [pscustomobject]@{
        Title = 'Build PowerPoint decks without coordinates'
        Intro = 'Slides are built from theme layouts. `-Layout` is optional: the layout is chosen from the slide content, and the first slide uses the title layout. Deck plans and the designer deck fold into `New-OfficePowerPoint` and `PptSlide`. `-X`, `-Y`, `-Width` and `-Height` remain for deliberate placement and take lengths such as `2cm`.'
        Old   = @('New-OfficePowerPoint', 'Add-OfficePowerPointSlide', 'Add-OfficePowerPointBullets', 'Add-OfficePowerPointTextBox', 'New-OfficePowerPointDeckPlan', 'Add-OfficePowerPointDesignerDeck')
    }
    [pscustomobject]@{
        Title  = 'Convert files'
        Intro  = 'File-to-file `ConvertTo-*` and `ConvertFrom-*` commands and `Export-OfficeDocumentPdf` become `Convert-OfficeDocument -Path -To -OutputPath`. Format-specific settings go through `-Options`. Batches run in parallel with `-ThrottleLimit`, and `-Resume` continues an interrupted batch. `ConvertTo-OfficeMarkdown` and `ConvertTo-OfficeCsv` stay, because they turn objects into text rather than converting files.'
        Old    = @('ConvertTo-OfficeDocument', 'Export-OfficeDocumentPdf', 'ConvertTo-OfficeWordMarkdown', 'ConvertFrom-OfficeWordMarkdown', 'ConvertTo-OfficePdfWord', 'ConvertTo-OfficeExcelHtml')
        Target = 'Convert-OfficeDocument'
    }
    [pscustomobject]@{
        Title  = 'Export pages, slides and sheets as images'
        Intro  = 'The per-format `Export-Office*Image` commands become `Export-OfficeDocumentImage`. Several images go to `-OutputDirectory`; `-OutputPath` writes exactly one file and is never treated as a folder.'
        Target = 'Export-OfficeDocumentImage'
        AllCmdletsForTarget = $true
    }
    [pscustomobject]@{
        Title = 'Read and edit PDF files'
        Intro = '`Get-OfficePdf` becomes `Open-OfficePdf` for editing. Reading text, structure and metadata uses the cross-format `Read-OfficeDocument` and `Get-OfficeDocumentInfo`, which work the same for every format.'
        Old   = @('Get-OfficePdf', 'Read-OfficePdf', 'Get-OfficePdfText', 'Get-OfficePdfInfo')
    }
)

#endregion

#region Helpers

$script:DataWarnings = [System.Collections.Generic.List[string]]::new()
function Add-DataWarning([string] $Message) { $script:DataWarnings.Add($Message) }

function Get-Field($Object, [string] $Name) {
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Test-Text($Value) { return ($Value -is [string]) -and -not [string]::IsNullOrWhiteSpace($Value) }

$jsonOptions = @{}
if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $jsonOptions.DateKind = 'String' }
function Read-MapFile([System.IO.FileInfo] $File) {
    try {
        return Get-Content -LiteralPath $File.FullName -Raw -Encoding utf8 | ConvertFrom-Json @jsonOptions
    } catch {
        throw "$($File.Name): invalid JSON. $($_.Exception.Message)"
    }
}

# Sorts by a list of string keys with ordinal comparison; the last key must make every item unique.
function Sort-Ordinal([object[]] $Items, [scriptblock] $Key) {
    if (-not $Items -or $Items.Count -eq 0) { return , @() }
    $keys = [System.Collections.Generic.List[string]]::new($Items.Count)
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $keys.Add([string](& $Key $Items[$i]) + [char]0 + $i.ToString('D6', [cultureinfo]::InvariantCulture))
    }
    $keys.Sort([StringComparer]::Ordinal)
    $sorted = [object[]]::new($Items.Count)
    for ($i = 0; $i -lt $keys.Count; $i++) {
        $sorted[$i] = $Items[[int]$keys[$i].Substring($keys[$i].LastIndexOf([char]0) + 1)]
    }
    return , $sorted
}

function Get-SortName([string] $Text) {
    if ($null -eq $Text) { $Text = '' }
    return $Text.ToUpperInvariant() + [char]1 + $Text
}

function ConvertTo-SingleLine([string] $Text) {
    if ($null -eq $Text) { return '' }
    return (($Text -replace '\s*[\r\n]+\s*', ' ').Trim())
}

# Inline code span that survives backticks in the content.
function Format-Code([string] $Text) {
    $Text = ConvertTo-SingleLine $Text
    if ($Text.Length -eq 0) { return '' }
    $longest = 0
    foreach ($match in [regex]::Matches($Text, '`+')) { $longest = [Math]::Max($longest, $match.Length) }
    $fence = '`' * ($longest + 1)
    if ($longest -gt 0 -or $Text.StartsWith(' ') -or $Text.EndsWith(' ')) { return "$fence $Text $fence" }
    return "$fence$Text$fence"
}

# Escapes text for Markdown inline content. Pipes are escaped later, per table cell.
function Format-PlainText([string] $Text) {
    $builder = [System.Text.StringBuilder]::new()
    foreach ($character in $Text.ToCharArray()) {
        switch ($character) {
            '\' { [void]$builder.Append('\\') }
            '*' { [void]$builder.Append('\*') }
            '_' { [void]$builder.Append('\_') }
            '`' { [void]$builder.Append('\`') }
            '[' { [void]$builder.Append('\[') }
            ']' { [void]$builder.Append('\]') }
            '<' { [void]$builder.Append('&lt;') }
            '>' { [void]$builder.Append('&gt;') }
            default { [void]$builder.Append($character) }
        }
    }
    return $builder.ToString()
}

# Prose from the map: command names, parameters and variables become code, the rest is escaped.
$CodeTokenPattern = '(?<![\w\-.$])(?:[A-Z][a-z]+-[A-Z][A-Za-z0-9*{}]*|-[A-Z][A-Za-z0-9]*\*?|\$[A-Za-z_]\w*)(?![\w\-])'
function Format-Prose([string] $Text) {
    $Text = ConvertTo-SingleLine $Text
    $builder = [System.Text.StringBuilder]::new()
    $position = 0
    foreach ($match in [regex]::Matches($Text, $CodeTokenPattern)) {
        [void]$builder.Append((Format-PlainText $Text.Substring($position, $match.Index - $position)))
        [void]$builder.Append((Format-Code $match.Value))
        $position = $match.Index + $match.Length
    }
    [void]$builder.Append((Format-PlainText $Text.Substring($position)))
    return $builder.ToString()
}

function Test-ProseHint([string] $Hint) {
    return $Hint -match '^[A-Z][a-z]+ [a-z]' -and $Hint -notmatch '^Use '
}

# Splits "code   (note)" hints into the code and the note.
function Split-HintNote([string] $Hint) {
    $match = [regex]::Match($Hint, '^(?<code>.*?\S)\s+\((?<note>[a-z][^()]*)\)\s*$')
    if ($match.Success) { return $match.Groups['code'].Value, $match.Groups['note'].Value }
    return $Hint, $null
}

function Format-Hint([string] $Hint) {
    $Hint = ConvertTo-SingleLine $Hint
    if ($Hint.Length -eq 0) { return '' }
    $use = [regex]::Match($Hint, '^Use (?<rest>.+?)\.?$')
    if ($use.Success) {
        $rest = $use.Groups['rest'].Value
        if ($rest -match '^[A-Z$][\w\-]*(\s|$)' -and $rest -notmatch ', | or | inside |^the ') {
            return 'Use ' + (Format-Code $rest)
        }
        return Format-Prose $Hint
    }
    if (Test-ProseHint $Hint) { return Format-Prose $Hint }
    $code, $note = Split-HintNote $Hint
    if ($note) { return (Format-Code $code) + ' (' + (Format-Prose $note) + ')' }
    return Format-Code $Hint
}

# Hint as a line of a PowerShell code block, or $null when the hint is prose.
function Get-HintCodeLine([string] $Hint) {
    $Hint = ConvertTo-SingleLine $Hint
    if ($Hint.Length -eq 0 -or $Hint -match '^Use ' -or (Test-ProseHint $Hint)) { return $null }
    $code, $note = Split-HintNote $Hint
    if ($note) { return "$code   # $note" }
    return $code
}

function Format-ParameterTarget([string] $Value) {
    if ($Value -match '^\w+$') { return Format-Code "-$Value" }
    if ($Value -match '^(?<name>\w+) \((?<note>.+)\)$') { return (Format-Code "-$($Matches.name)") + ' (' + (Format-Prose $Matches.note) + ')' }
    if ($Value -match '^\w+\.\w+$') { return Format-Code $Value }
    if ($Value -match '^\w+ / \w+$') { return (($Value -split ' / ') | ForEach-Object { Format-Code "-$_" }) -join ' / ' }
    if ($Value -match ' \| ') { return (($Value -split ' \| ') | ForEach-Object { Format-ParameterTarget $_ }) -join ' or ' }
    if ($Value -match '^[A-Z][a-z]+-[A-Z]') { return Format-Code $Value }
    if ($Value -match '^\w+ ') { return Format-Code "-$Value" }
    return Format-Code $Value
}

# Renders a "parameters" object as a list of changes, grouping old names that share a target.
function Get-ParameterChanges($Parameters) {
    $groups = [ordered]@{}
    $removed = [System.Collections.Generic.List[string]]::new()
    $changed = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $Parameters) { return , @() }
    foreach ($property in $Parameters.PSObject.Properties) {
        $old = $property.Name
        $new = $property.Value
        if ($null -eq $new) { $removed.Add($old); continue }
        $new = [string]$new
        if ($new -eq $old) { $changed.Add($old); continue }
        if (-not $groups.Contains($new)) { $groups[$new] = [System.Collections.Generic.List[string]]::new() }
        $groups[$new].Add($old)
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($target in $groups.Keys) {
        $olds = ($groups[$target] | ForEach-Object { Format-Code "-$_" }) -join ', '
        $lines.Add("$olds &rarr; $(Format-ParameterTarget $target)")
    }
    if ($changed.Count -gt 0) {
        $lines.Add((($changed | ForEach-Object { Format-Code "-$_" }) -join ', ') + ' (same name, changed)')
    }
    if ($removed.Count -gt 0) {
        $lines.Add('Removed: ' + (($removed | ForEach-Object { Format-Code "-$_" }) -join ', '))
    }
    return , $lines.ToArray()
}

function Format-Cell([string] $Text) {
    if ($null -eq $Text -or $Text.Length -eq 0) { return ' ' }
    return $Text.Replace('|', '\|')
}

function Get-CodeFence([string] $Code) {
    $longest = 2
    foreach ($match in [regex]::Matches($Code, '(?m)^\s*(`{3,})')) { $longest = [Math]::Max($longest, $match.Groups[1].Length) }
    return '`' * ($longest + 1)
}

$script:Lines = [System.Collections.Generic.List[string]]::new()
function Add-Line([string] $Text = '') { $script:Lines.Add($Text) }
function Add-Heading([int] $Depth, [string] $Text) {
    Add-Line (('#' * ($HeadingLevel + $Depth)) + ' ' + $Text)
    Add-Line
}
function Add-Paragraph([string] $Text) {
    Add-Line $Text
    Add-Line
}
function Add-Table([string[]] $Header, [object[]] $Rows) {
    # A single row arrives unrolled by the pipeline; wrap it again.
    if ($Rows -and $Rows.Count -gt 0 -and $Rows[0] -isnot [array]) { $Rows = , $Rows }
    Add-Line ('| ' + ($Header -join ' | ') + ' |')
    Add-Line ('|' + ((@($Header) | ForEach-Object { ' --- ' }) -join '|') + '|')
    foreach ($row in $Rows) {
        Add-Line ('| ' + ((@($row) | ForEach-Object { Format-Cell $_ }) -join ' | ') + ' |')
    }
    Add-Line
}
function Add-CodeBlock([string] $Language, [string[]] $CodeLines) {
    $code = $CodeLines -join "`n"
    $fence = Get-CodeFence $code
    Add-Line "$fence$Language"
    foreach ($line in $CodeLines) { Add-Line $line }
    Add-Line $fence
    Add-Line
}
function Add-DetailsStart([string] $Summary) {
    Add-Line '<details>'
    Add-Line "<summary>$Summary</summary>"
    Add-Line
}
function Add-DetailsEnd {
    Add-Line '</details>'
    Add-Line
}

function Get-Plural([int] $Count, [string] $Singular, [string] $Plural = "${Singular}s") {
    if ($Count -eq 1) { return "1 $Singular" }
    return "$Count $Plural"
}

#endregion

#region Load the map

$powerShellDirectory = Join-Path $MapDirectory 'powershell'
$csharpDirectory = Join-Path $MapDirectory 'csharp'
foreach ($directory in $powerShellDirectory, $csharpDirectory) {
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { throw "Map directory not found: $directory" }
}

$psFiles = @(Get-ChildItem -LiteralPath $powerShellDirectory -Filter '*.json' -File)
$csFiles = @(Get-ChildItem -LiteralPath $csharpDirectory -Filter '*.json' -File)
if ($psFiles.Count -eq 0 -or $csFiles.Count -eq 0) { throw "No map files under $MapDirectory." }

$areas = [System.Collections.Generic.List[object]]::new()
foreach ($file in $psFiles) {
    $data = Read-MapFile $file
    $areaName = Get-Field $data 'area'
    if (-not (Test-Text $areaName)) { $areaName = $file.BaseName }
    $entries = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($entry in @(Get-Field $data 'map')) {
        if ($null -eq $entry) { continue }
        $entries.Add([pscustomobject]@{
                Id    = "powershell/$($file.Name)#$index"
                File  = "powershell/$($file.Name)"
                Index = $index
                Area  = $areaName
                Entry = [pscustomobject]@{
                    old          = Get-Field $entry 'old'
                    kind         = Get-Field $entry 'kind'
                    action       = Get-Field $entry 'action'
                    new          = Get-Field $entry 'new'
                    parameterSet = Get-Field $entry 'parameterSet'
                    parameters   = Get-Field $entry 'parameters'
                    hint         = Get-Field $entry 'hint'
                    reason       = Get-Field $entry 'reason'
                }
            })
        $index++
    }
    $areas.Add([pscustomobject]@{
            Name     = $areaName
            File     = "powershell/$($file.Name)"
            Commands = @(Get-Field $data 'commands' | Where-Object { $null -ne $_ })
            Entries  = $entries.ToArray()
        })
}
$areaRank = { param($name) $i = [Array]::IndexOf([string[]]$PowerShellAreaOrder, $name); if ($i -lt 0) { $i = 999 }; '{0:D3}' -f $i }
$areas = Sort-Ordinal $areas.ToArray() { param($a) (& $areaRank $a.Name) + [char]1 + (Get-SortName $a.Name) }

$packages = [System.Collections.Generic.List[object]]::new()
foreach ($file in $csFiles) {
    $data = Read-MapFile $file
    $packageName = Get-Field $data 'package'
    if (-not (Test-Text $packageName)) { $packageName = $file.BaseName }
    $changes = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($change in @(Get-Field $data 'changes')) {
        if ($null -eq $change) { continue }
        $changes.Add([pscustomobject]@{
                Id      = "csharp/$($file.Name)#$index"
                File    = "csharp/$($file.Name)"
                Index   = $index
                Package = $packageName
                Entry   = [pscustomobject]@{
                    old         = Get-Field $change 'old'
                    kind        = Get-Field $change 'kind'
                    action      = Get-Field $change 'action'
                    new         = Get-Field $change 'new'
                    replacement = Get-Field $change 'replacement'
                    reason      = Get-Field $change 'reason'
                    category    = Get-Field $change 'category'
                }
            })
        $index++
    }
    $packages.Add([pscustomobject]@{
            Name         = $packageName
            File         = "csharp/$($file.Name)"
            Data         = $data
            Changes      = $changes.ToArray()
            PackageMoves = @(Get-Field $data 'packageMoves' | Where-Object { $null -ne $_ })
        })
}
$packageRank = { param($name) $i = [Array]::IndexOf([string[]]$CSharpPackageOrder, $name); if ($i -lt 0) { $i = 999 }; '{0:D3}' -f $i }
$packages = Sort-Ordinal $packages.ToArray() { param($p) (& $packageRank $p.Name) + [char]1 + (Get-SortName $p.Name) }

$allPsEntries = @($areas | ForEach-Object { $_.Entries })
$allCsEntries = @($packages | ForEach-Object { $_.Changes })
$allCommands = @($areas | ForEach-Object { $_.Commands })

$targetCommandNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$targetAliasNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$commandByName = @{}
foreach ($command in $allCommands) {
    $name = Get-Field $command 'name'
    if (-not (Test-Text $name)) { continue }
    [void]$targetCommandNames.Add($name)
    $commandByName[$name] = $command
    $alias = Get-Field $command 'alias'
    foreach ($a in @($alias)) { if (Test-Text $a) { [void]$targetAliasNames.Add($a) } }
}

#endregion

#region Map data checks (warnings)

$seenOld = @{}
foreach ($item in $allPsEntries) {
    $e = $item.Entry
    $where = "$($item.File) '$(Get-Field $e 'old')'"
    $old = Get-Field $e 'old'
    $action = Get-Field $e 'action'
    $kind = Get-Field $e 'kind'
    $new = Get-Field $e 'new'
    $hint = Get-Field $e 'hint'
    if (-not (Test-Text $old)) { Add-DataWarning "$($item.Id): entry has no 'old' name."; continue }
    if ($seenOld.ContainsKey($old)) { Add-DataWarning "$where is also mapped in $($seenOld[$old])." } else { $seenOld[$old] = $item.File }
    if (-not $PowerShellActionLabels.ContainsKey([string]$action)) { Add-DataWarning "$where has unknown action '$action'." }
    if ($kind -notin 'cmdlet', 'alias') { Add-DataWarning "$where has unknown kind '$kind'." }
    if ($action -ne 'keep' -and -not (Test-Text (Get-Field $e 'reason'))) { Add-DataWarning "$where ($action) has no reason." }
    if ($action -notin 'keep' -and -not (Test-Text $hint)) { Add-DataWarning "$where ($action) has no hint." }
    if ($action -eq 'remove' -and $null -ne $new) { Add-DataWarning "$where is removed but names new '$new'." }
    if ($action -ne 'remove' -and -not (Test-Text $new)) { Add-DataWarning "$where ($action) has no new name." }
    if ($action -ne 'remove' -and (Test-Text $new) -and -not $targetCommandNames.Contains($new) -and -not $targetAliasNames.Contains($new)) {
        Add-DataWarning "$where maps to '$new', which is not a target command or alias."
    }
    if (Test-Text $hint) {
        foreach ($match in [regex]::Matches($hint, '\b[A-Z][a-z]+-Office[A-Za-z]+\b')) {
            if (-not $targetCommandNames.Contains($match.Value)) { Add-DataWarning "$where hint names '$($match.Value)', which is not a target command." }
        }
        foreach ($match in [regex]::Matches($hint, '(?<=-(?:To|From) )[A-Za-z]+')) {
            if ($match.Value -notin $ConversionFormatNames) { Add-DataWarning "$where hint uses format '$($match.Value)'; format names are $($ConversionFormatNames -join ', ')." }
        }
        foreach ($match in [regex]::Matches($hint, '(?<=-Layout )[A-Za-z]+')) {
            if ($match.Value -notin $SlideLayouts -and $hint -match 'Ppt|PowerPoint') { Add-DataWarning "$where hint uses slide layout '$($match.Value)', which is not one of the designed layouts." }
        }
        foreach ($match in [regex]::Matches($hint, '(?<=-OutputPath )\S+')) {
            if ($match.Value -notmatch '\.[A-Za-z0-9]+$' -and $match.Value -notmatch '^\$') { Add-DataWarning "$where hint passes '$($match.Value)' to -OutputPath, which looks like a folder (-OutputDirectory)." }
        }
    }
    $parameters = Get-Field $e 'parameters'
    if ($null -ne $parameters) {
        foreach ($property in $parameters.PSObject.Properties) {
            $value = [string]$property.Value
            $command = [regex]::Match($value, '^[A-Z][a-z]+-Office[A-Za-z]+')
            if ($command.Success -and -not $targetCommandNames.Contains($command.Value)) {
                Add-DataWarning "$where moves -$($property.Name) to '$($command.Value)', which is not a target command."
            }
        }
    }
}

# A parameter renamed in one entry but still present on a target command under its old name.
$renames = @{}
foreach ($item in $allPsEntries) {
    $parameters = Get-Field $item.Entry 'parameters'
    if ($null -eq $parameters) { continue }
    foreach ($property in $parameters.PSObject.Properties) {
        if ($property.Value -is [string] -and $property.Value -match '^\w+$' -and $property.Value -ne $property.Name) {
            if (-not $renames.ContainsKey($property.Name)) { $renames[$property.Name] = [System.Collections.Generic.List[string]]::new() }
            $renames[$property.Name].Add("$($item.Entry.old) -> -$($property.Value)")
        }
    }
}
foreach ($command in (Sort-Ordinal $allCommands { param($c) Get-SortName (Get-Field $c 'name') })) {
    $names = @(foreach ($set in @(Get-Field $command 'parameterSets')) { foreach ($p in @(Get-Field $set 'parameters')) { Get-Field $p 'name' } }) | Where-Object { $_ } | Sort-Object -Unique
    foreach ($name in $names) {
        if ($renames.ContainsKey($name) -and $name -in 'TemplatePath', 'FilePath', 'InputPath', 'SourcePath', 'SheetName', 'SheetIndex', 'FontSize', 'FontName', 'Strike') {
            Add-DataWarning "Target command $(Get-Field $command 'name') keeps -$name, which the map renames elsewhere ($($renames[$name][0]))."
        }
    }
}

foreach ($item in $allCsEntries) {
    $e = $item.Entry
    $old = [string](Get-Field $e 'old')
    $new = [string](Get-Field $e 'new')
    $action = Get-Field $e 'action'
    $short = if ($old.Length -gt 90) { $old.Substring(0, 90) + '...' } else { $old }
    $where = "$($item.File) '$short'"
    if (-not $CSharpActionLabels.ContainsKey([string]$action)) { Add-DataWarning "$where has unknown action '$action'." }
    if (-not (Test-Text (Get-Field $e 'reason')) -and -not (Test-Text (Get-Field $e 'replacement'))) { Add-DataWarning "$where has neither reason nor replacement." }
    if ($action -notin 'remove', 'internalize' -and -not (Test-Text $new)) { Add-DataWarning "$where ($action) has no new member." }
    if ($new -match '\b\w*Document\.From(Html|Markdown|Pdf|Word|Excel|PowerPoint|Rtf)\w*\(') { Add-DataWarning "$where maps to static '$($Matches[0])...'; the design has no static From... conversion factories (use To{Target} on the source)." }
    if ($new -match '\(or ') { Add-DataWarning "$where lists an unresolved alternative in 'new'." }
    if ((Get-Field $e 'reason') -match 'From\{Source\}') { Add-DataWarning "$where reason cites {Format}Document.From{Source}, which the design dropped." }
    if ($new -match '\bOfficeFormat\b') { Add-DataWarning "$where uses OfficeFormat; the design names formats with the existing OfficeDocumentFormat enum." }
    if ("$old $new" -match '`\d') { Add-DataWarning "$where uses a reflection generic name (``1) instead of <T>." }
}
foreach ($package in $packages) {
    foreach ($move in $package.PackageMoves) {
        if ((Get-Field $move 'to') -match '\(or ') { Add-DataWarning "$($package.File) packageMoves '$(Get-Field $move 'from') -> $(Get-Field $move 'to')' is unresolved." }
    }
    $project = Get-Field $package.Data 'project'
    if ($package.Name -eq 'OfficeIMO' -and $null -ne $project) {
        $tfms = @(Get-Field $project 'targetFrameworks')
        if (($tfms -join ';') -ne ($FrontDoorFrameworks -join ';')) { Add-DataWarning "$($package.File) project.targetFrameworks is '$($tfms -join ', ')'; the design says the OfficeIMO package targets $($FrontDoorFrameworks -join ', ')." }
    }
}

#endregion

#region Counts

$oldCmdlets = @($allPsEntries | Where-Object { $_.Entry.kind -eq 'cmdlet' })
$oldAliases = @($allPsEntries | Where-Object { $_.Entry.kind -eq 'alias' })
$targetAliasCount = $targetAliasNames.Count
$removedCmdlets = @($oldCmdlets | Where-Object { $_.Entry.action -eq 'remove' })
$invoiceEntries = @($allPsEntries | Where-Object { $_.Entry.action -eq 'remove' -and ([string]$_.Entry.old -match 'Invoice') })

#endregion

#region Section 1: what changed

$emitted = @{}
function Register-Row($Item) {
    if ($emitted.ContainsKey($Item.Id)) { $emitted[$Item.Id]++ } else { $emitted[$Item.Id] = 1 }
}

Add-Heading 0 'OfficeIMO and PSWriteOffice 3.6'
Add-Paragraph ('OfficeIMO 3.6 and PSWriteOffice 3.6 change the public API in both languages and are released together. ' +
    'Nothing keeps a compatibility wrapper, so update scripts and code with the rules and tables below. ' +
    "PSWriteOffice goes from $($oldCmdlets.Count) commands and $($oldAliases.Count) aliases to $($targetCommandNames.Count) commands and $targetAliasCount aliases; " +
    'every job that worked before still works, usually through fewer commands. The [API design](Docs/officeimo.api-design.md) explains the rules.')

Add-Heading 1 'What changed'
Add-Paragraph 'For PowerShell scripts:'
@(
    '- **Lifecycle.** Every authoring format has `New-Office{Format}`, `Open-Office{Format}`, `Save-Office{Format}` and `Close-Office{Format}`. With a content block, `New-` and `Open-` save and close the file when the block ends; if the block throws, nothing is saved. Without a block they return the live document. `-NoSave` is removed, `Get-Office{Format}` commands that opened files become `Open-`, and `-Open` becomes `-Show`. `New-` refuses to replace an existing file unless you add `-Force`.'
    '- **Paths.** `-Path` is the input file when a command reads one (with wildcards), otherwise the file it creates. `-LiteralPath` takes input without wildcard expansion and binds from `Get-ChildItem`. `-OutputPath` is always one output file and never silently becomes a folder; commands that write several files take `-OutputDirectory`. `Export-` and `Convert-` replace an existing output unless you add `-NoClobber`.'
    '- **One name per idea.** `-Document` replaces `-Presentation`, `-Workbook` and similar. `-Sheet` takes a name, a 1-based position or a sheet and replaces `-SheetName`/`-SheetIndex`; select a sheet named `1` with `-Sheet ''1''`. Charts read `-Category` and `-Value` instead of `-CategoryProperty` and `-SeriesProperty`.'
    '- **Switches instead of `$true`.** On/off parameters are switches: write `-KeepWithNext`, not `-KeepWithNext $true`. `-Bold:$false` still turns formatting off explicitly, and defaults that are on are turned off with `-No{Feature}` switches such as `-NoFreeze` and `-NoAutoFit`.'
    '- **Lengths.** A bare number means points; a unit is written directly: `-Width 5cm`, `-SpacingBefore 12pt`, `-Margin 1in`. Parameter names no longer carry units, so `-WidthCm`, `-HeightPixels` and `-SpacingBeforePoints` become `-Width`, `-Height` and `-SpacingBefore`. Excel column widths stay in characters.'
    '- **Positions are 1-based** everywhere: pages, slides, sheets, sections, table rows and columns, list items and insert positions. Scripts that passed 0-based indexes must add one.'
    '- **Passwords.** `-Password` is a `SecureString`. It also accepts a plain string, with a warning, or a `PSCredential`, whose password is used. There is no separate `-Credential` parameter.'
    '- **Themes by default.** New documents start from the `Modern` theme; choose another with `-Theme` (a built-in name, a theme object or a template path). Opening an existing document never restyles it.'
    '- **Loss handling.** Converting to another format reports what could not be carried and continues; add `-FailOnLoss` to make loss an error. Saving a document in its own format fails if content would be lost; add `-AllowLoss` to accept it.'
    '- **Options.** Option objects come from `New-Office{Format}{Purpose}Options`, and every `-Options` parameter also accepts a hashtable of option property names. Unknown keys are errors.'
    '- **Content blocks.** Styled text is written with child commands (`WordText`, `PptText`, `PdfText`, `ExcelText`); `-Run` hashtables of parallel arrays are removed. Each DSL command has exactly one alias, and `Ppt` is the only PowerPoint prefix. Plural, spelling, verb-last and unprefixed alias twins are removed.'
    '- **`-WhatIf` and `-Confirm`** cover file-system effects: `New`, `Save`, `Export`, `Convert`, `Repair`, `Protect` and `Remove` commands.'
    ("- **Electronic invoices** are no longer part of PSWriteOffice ($(Get-Plural $invoiceEntries.Count 'command and alias' 'commands and aliases') removed). Use the `OfficeIMO.Invoicing` packages from .NET.")
    '- **Removed-command hints.** When a script calls a command or alias that no longer exists, PSWriteOffice prints its replacement, taken from the reference tables below.'
) | ForEach-Object { Add-Line $_ }
Add-Line
Add-Paragraph 'For .NET code:'
@(
    '- **Lifecycle.** Every document type, including `PdfDocument`, is created with `Create`, opened with `Load` or `LoadAsync`, saved with `Save` or `SaveAsync`, and disposed with `using`. Options are `{Format}CreateOptions`, `{Format}LoadOptions` and `{Format}SaveOptions`.'
    '- **One authoring model.** `AsFluent()`, `WordFluentDocument`, `ExcelFluentWorkbook` and their builders are removed; use the object model. `Add*` methods return what they create, and `AddImage` is the only image verb. Excel''s `SheetComposer` stays, reached through `sheet.Report()`.'
    '- **Conversions** are called on the source: `SaveAs{Target}(path or stream)` writes the file and returns `OfficeConversionReport`; `To{Target}()` returns the target model. The throwing and `...Result` pairs merge. `OfficeConverter` is the entry point for file-to-file and batch work.'
    '- **Units and colors.** Geometry takes `OfficeLength` (`OfficeLength.Centimeters(2.5)`, `OfficeLength.Parse("12pt")`) instead of raw EMUs, twips or unit-suffixed parameters. `OfficeColor` is the only color type and moves to the `OfficeIMO` namespace.'
    '- **Packages.** Format packages no longer carry other format engines, and `OfficeIMO.Word.Pdf` and `OfficeIMO.PowerPoint.Pdf` change direction. See [.NET packages](#net-packages).'
) | ForEach-Object { Add-Line $_ }
Add-Line

#endregion

#region Section 2: common jobs

$entryByOld = @{}
foreach ($item in $allPsEntries) {
    $old = Get-Field $item.Entry 'old'
    if ((Test-Text $old) -and -not $entryByOld.ContainsKey($old)) { $entryByOld[$old] = $item }
}

Add-Heading 1 'Common jobs'
Add-Paragraph 'Each job lists the commands a typical script uses, what they become and the parameters that change, followed by the new form.'

foreach ($job in $CommonJobs) {
    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($old in @(Get-Field $job 'Old')) {
        if (-not $old) { continue }
        if ($entryByOld.ContainsKey($old)) { $items.Add($entryByOld[$old]) } else { Add-DataWarning "Common job '$($job.Title)' expects '$old', which is not in the map." }
    }
    $target = Get-Field $job 'Target'
    $others = @()
    if ($target) {
        $forTarget = @($allPsEntries | Where-Object { $_.Entry.new -eq $target -and $_.Entry.action -ne 'keep' })
        if (Get-Field $job 'AllCmdletsForTarget') {
            foreach ($candidate in (Sort-Ordinal @($forTarget | Where-Object { $_.Entry.kind -eq 'cmdlet' }) { param($i) (Get-SortName $i.Entry.old) + [char]1 + $i.Id })) {
                if (-not $items.Contains($candidate)) { $items.Add($candidate) }
            }
        }
        $others = @($forTarget | Where-Object { -not $items.Contains($_) })
    }
    if ($items.Count -eq 0) { continue }

    Add-Heading 2 $job.Title
    Add-Paragraph $job.Intro
    $rows = foreach ($item in $items) {
        $e = $item.Entry
        $newText = if (Test-Text $e.new) { Format-Code $e.new } else { 'Removed' }
        , @((Format-Code $e.old), $newText, ((Get-ParameterChanges (Get-Field $e 'parameters')) -join '<br>'))
    }
    Add-Table @('3.5', '3.6', 'Parameter changes') $rows

    $code = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $items) {
        $line = Get-HintCodeLine (Get-Field $item.Entry 'hint')
        if ($line -and -not $code.Contains($line)) { $code.Add($line) }
    }
    $targets = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $items) { $n = Get-Field $item.Entry 'new'; if ((Test-Text $n) -and -not $targets.Contains($n)) { $targets.Add($n) } }
    # The realistic example (the second one) of the first targets, at most $MaxJobExamples per job.
    $shown = 0
    foreach ($name in $targets) {
        if ($shown -ge $MaxJobExamples) { break }
        $command = $commandByName[$name]
        $examples = @(Get-Field $command 'examples' | Where-Object { $null -ne $_ -and (Test-Text (Get-Field $_ 'code')) })
        if ($examples.Count -eq 0) { continue }
        $example = $examples[[Math]::Min(1, $examples.Count - 1)]
        $shown++
        & {
            if ($code.Count -gt 0) { $code.Add('') }
            $title = Get-Field $example 'title'
            if (Test-Text $title) { $code.Add("# $(ConvertTo-SingleLine $title)") }
            foreach ($line in ((Get-Field $example 'code') -split '\r?\n')) { $code.Add($line.TrimEnd()) }
        }
    }
    if ($code.Count -gt 0) { Add-CodeBlock 'powershell' $code.ToArray() }

    if ($others.Count -gt 0) {
        $otherCmdlets = @($others | Where-Object { $_.Entry.kind -eq 'cmdlet' }).Count
        $otherAliases = $others.Count - $otherCmdlets
        $parts = @()
        if ($otherCmdlets) { $parts += Get-Plural $otherCmdlets 'more command' 'more commands' }
        if ($otherAliases) { $parts += Get-Plural $otherAliases 'alias' 'aliases' }
        Add-Paragraph "$($parts -join ' and ') also map to $(Format-Code $target); the [reference tables](#powershell-reference) list each one."
    }
}

#endregion

#region Section 3: .NET packages

Add-Heading 1 '.NET packages'

$allMoves = [System.Collections.Generic.List[object]]::new()
foreach ($package in $packages) {
    $i = 0
    foreach ($move in $package.PackageMoves) {
        $allMoves.Add([pscustomobject]@{ Owner = $package.Name; Index = $i; From = [string](Get-Field $move 'from'); To = [string](Get-Field $move 'to'); Scope = [string](Get-Field $move 'scope'); Reason = [string](Get-Field $move 'reason') })
        $i++
    }
}
function Get-MoveTarget([string] $To) {
    $value = $To -replace '^\(new\)\s*', ''
    $value = $value -replace '\s*\(.*$', ''
    return $value.Trim()
}
$relocations = @($allMoves | Where-Object { $_.To -notmatch '\(none\)|\(stays\)|\(remove references\)|\(removal\)' -and (Get-MoveTarget $_.To) })
$dependencyRemovals = @($allMoves | Where-Object { $_.To -match '\(remove references\)|\(removal\)' })

# Keep one record per source and target, preferring the one in the source package's own file.
$uniqueMoves = [System.Collections.Generic.List[object]]::new()
foreach ($group in ($relocations | Group-Object { "$($_.From)|$(Get-MoveTarget $_.To)" })) {
    $own = @($group.Group | Where-Object { $_.Owner -eq $_.From })
    foreach ($move in $(if ($own.Count) { $own } else { $group.Group })) { $uniqueMoves.Add($move) }
}
$moveSort = { param($m) (& $packageRank $m.From) + [char]1 + (Get-SortName $m.From) + [char]1 + (Get-SortName (Get-MoveTarget $m.To)) + [char]1 + (& $packageRank $m.Owner) + [char]1 + ('{0:D4}' -f $m.Index) }
$uniqueMoves = Sort-Ordinal $uniqueMoves.ToArray() $moveSort

$flipped = @()
foreach ($move in $uniqueMoves) {
    $source = $move.From
    $incoming = @($uniqueMoves | Where-Object { (Get-MoveTarget $_.To) -eq $source })
    if ($incoming.Count -gt 0 -and $flipped -notcontains $source) { $flipped += $source }
}
$flipped = Sort-Ordinal $flipped { param($p) (& $packageRank $p) + [char]1 + (Get-SortName $p) }

if ($flipped.Count -gt 0) {
    $names = ($flipped | ForEach-Object { Format-Code $_ }) -join ' and '
    $verb = if ($flipped.Count -eq 1) { 'changes' } else { 'change' }
    Add-Heading 2 "$(($flipped | ForEach-Object { $_ }) -join ' and ') $verb direction"
    Add-Paragraph "$names keep their package IDs but change direction. In 3.5 they convert PDF into the Office format; in 3.6 they export the Office format to PDF. Updating the version without changing code therefore gives a project a different API under the same package ID."
    $rows = foreach ($name in $flipped) {
        $out = @($uniqueMoves | Where-Object { $_.From -eq $name })
        $in = @($uniqueMoves | Where-Object { (Get-MoveTarget $_.To) -eq $name })
        $before = ($out | ForEach-Object { "Moves to $(Format-Code (Get-MoveTarget $_.To)): $(Format-Prose $_.Scope)" }) -join '<br>'
        $after = ($in | ForEach-Object { "From $(Format-Code $_.From): $(Format-Prose $_.Scope)" }) -join '<br>'
        , @((Format-Code $name), $before, $after)
    }
    Add-Table @('Package', '3.5 content', '3.6 content') $rows
    foreach ($name in $flipped) {
        $reverse = @($uniqueMoves | Where-Object { $_.From -eq $name } | ForEach-Object { Get-MoveTarget $_.To }) | Select-Object -First 1
        $sourcePackage = @($uniqueMoves | Where-Object { (Get-MoveTarget $_.To) -eq $name } | ForEach-Object { $_.From }) | Select-Object -First 1
        Add-Line "- If you reference $(Format-Code $name) to read PDF files into $(Format-Code $sourcePackage) documents, reference $(Format-Code $reverse) instead. If you export $(Format-Code $sourcePackage) documents to PDF, add $(Format-Code $name); the export no longer comes with $(Format-Code $sourcePackage)."
    }
    Add-Line
}

$frontDoor = $packages | Where-Object Name -eq 'OfficeIMO' | Select-Object -First 1
Add-Heading 2 'The OfficeIMO package'
$frontDoorText = 'The `OfficeIMO` package is the front door for applications that want the common formats and conversions in one reference: `dotnet add package OfficeIMO`.'
$frontDoorText += " It targets $(($FrontDoorFrameworks | ForEach-Object { Format-Code $_ }) -join ', '); .NET Framework applications reference the format and adapter packages directly."
Add-Paragraph $frontDoorText
if ($frontDoor) {
    $references = @(Get-Field (Get-Field $frontDoor.Data 'project') 'projectReferences') | Where-Object { Test-Text $_ } | ForEach-Object { ($_ -replace '\s*\(.*$', '').Trim() } | Select-Object -Unique
    if ($references) {
        Add-Paragraph ('It references ' + (($references | ForEach-Object { Format-Code $_ }) -join ', ') + '.')
    }
}
Add-Paragraph '`OfficeConverter`, the batch types and the shared foundations, including `OfficeColor`, live in the `OfficeIMO` namespace. Format types keep their per-format namespaces.'

Add-Heading 2 'Package moves'
Add-Paragraph 'Format packages depend only on `OfficeIMO.Core` and their own third-party libraries. Conversions live in `OfficeIMO.{Source}.{Target}` adapter packages, and reconstruction from PDF lives in `OfficeIMO.Pdf.{Target}`.'
$rows = foreach ($move in $uniqueMoves) {
    $to = Get-MoveTarget $move.To
    $label = if ($move.To -match '^\(new\)') { "$(Format-Code $to) (new)" } else { Format-Code $to }
    , @((Format-Code $move.From), $label, (Format-Prose $move.Scope), (Format-Prose $move.Reason))
}
Add-Table @('From', 'To', 'What moves', 'Why') $rows

if ($dependencyRemovals.Count -gt 0) {
    Add-Heading 2 'Removed dependencies'
    Add-Paragraph 'These packages stop bringing another engine. Reference the adapter that you use directly.'
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $rows = foreach ($move in (Sort-Ordinal $dependencyRemovals $moveSort)) {
        $dropped = @([regex]::Matches($move.Scope, 'OfficeIMO(?:\.(?!csproj\b)[A-Za-z]+)*') | ForEach-Object Value | Where-Object { $_ -ne $move.From } | Select-Object -Unique)
        $key = "$($move.From)|$($dropped -join ',')"
        if (-not $seen.Add($key)) { continue }
        , @((Format-Code $move.From), (($dropped | ForEach-Object { Format-Code $_ }) -join ', '), (Format-Prose $move.Reason))
    }
    Add-Table @('Package', 'No longer references', 'Why') $rows
}

Add-Heading 2 'Target frameworks'
$frameworkRows = [System.Collections.Generic.List[object]]::new()
foreach ($package in $packages) {
    $frameworks = Get-Field $package.Data 'frameworks'
    if ($null -eq $frameworks) { continue }
    $current = @(@(Get-Field $frameworks 'current') | ForEach-Object { "$_" -split ';' } | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notmatch '^\$\(' } | Select-Object -Unique)
    $targetFrameworks = @(@(Get-Field $frameworks 'target') | ForEach-Object { "$_" -split ';' } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ($current.Count -eq 0 -and $targetFrameworks.Count -eq 0) { continue }
    $frameworkRows.Add(@((Format-Code $package.Name), (($current | ForEach-Object { Format-Code $_ }) -join ', '), (($targetFrameworks | ForEach-Object { Format-Code $_ }) -join ', ')))
}
Add-Paragraph 'Libraries target `netstandard2.0`, which serves .NET Framework 4.7.2 and later, and `net10.0`, identically on every build host. `net8.0` is dropped because .NET 8 reaches end of support in November 2026, and the separate `net472` build is dropped: .NET Framework applications use the `netstandard2.0` build.'
if ($frameworkRows.Count -gt 0) { Add-Table @('Package', '3.5', '3.6') $frameworkRows.ToArray() }

#endregion

#region Section 4: reference tables

Add-Heading 1 'PowerShell reference'
Add-Paragraph 'Every PSWriteOffice 3.5 command and alias appears once, sorted by name. **Use instead** is the hint PSWriteOffice prints when a script calls a command or alias that changed.'

foreach ($area in $areas) {
    $title = $PowerShellAreaTitles[$area.Name]
    if (-not $title) { $title = (Get-Culture).TextInfo.ToTitleCase($area.Name) }
    if ($area.Name -in 'markup', 'other' -or -not $PowerShellAreaTitles.ContainsKey($area.Name)) {
        $formats = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $area.Entries) {
            $m = [regex]::Match([string]$item.Entry.old, '^[A-Za-z]+-Office(?<noun>[A-Z]\w*)')
            if (-not $m.Success) { continue }
            foreach ($format in $FormatNames.Keys) {
                if ($m.Groups['noun'].Value.StartsWith($format, [StringComparison]::Ordinal)) {
                    if (-not $formats.Contains($FormatNames[$format])) { $formats.Add($FormatNames[$format]) }
                    break
                }
            }
        }
        if ($formats.Count -gt 0) {
            $sorted = Sort-Ordinal $formats.ToArray() { param($f) Get-SortName $f }
            $title = "$title ($($sorted -join ', '))"
        }
    }
    Add-Heading 2 $title

    $cmdlets = Sort-Ordinal @($area.Entries | Where-Object { $_.Entry.kind -ne 'alias' }) { param($i) (Get-SortName $i.Entry.old) + [char]1 + $i.Id }
    $aliases = Sort-Ordinal @($area.Entries | Where-Object { $_.Entry.kind -eq 'alias' }) { param($i) (Get-SortName $i.Entry.old) + [char]1 + $i.Id }

    $useInstead = {
        param($e)
        $parts = @()
        if (Test-Text (Get-Field $e 'hint')) { $parts += Format-Hint $e.hint }
        if ($e.action -eq 'remove' -and (Test-Text (Get-Field $e 'reason')) -and (Get-Field $e 'reason') -ne (Get-Field $e 'hint')) { $parts += Format-Prose $e.reason }
        if ($parts.Count -eq 0 -and $e.action -ne 'keep' -and (Test-Text (Get-Field $e 'reason'))) { $parts += Format-Prose $e.reason }
        $parts -join '<br>'
    }
    $newCell = {
        param($e)
        if (-not (Test-Text (Get-Field $e 'new'))) { return '&mdash;' }
        $text = Format-Code $e.new
        $set = Get-Field $e 'parameterSet'
        if (Test-Text $set) { $text += " ($(Format-Code $set) set)" }
        $text
    }

    if ($cmdlets.Count -gt 0) {
        $changedCount = @($cmdlets | Where-Object { $_.Entry.action -ne 'keep' }).Count
        Add-DetailsStart "$(Get-Plural $cmdlets.Count 'command'), $changedCount renamed, merged or removed"
        $rows = foreach ($item in $cmdlets) {
            Register-Row $item
            $e = $item.Entry
            $label = $PowerShellActionLabels[[string]$e.action]
            if (-not $label) { $label = [string]$e.action }
            , @((Format-Code $e.old), (& $newCell $e), $label, ((Get-ParameterChanges (Get-Field $e 'parameters')) -join '<br>'), (& $useInstead $e))
        }
        Add-Table @('3.5 command', '3.6 command', 'Change', 'Parameter changes', 'Use instead') $rows
        Add-DetailsEnd
    }
    if ($aliases.Count -gt 0) {
        $changedCount = @($aliases | Where-Object { $_.Entry.action -ne 'keep' }).Count
        Add-DetailsStart "$(Get-Plural $aliases.Count 'alias' 'aliases'), $changedCount changed or removed"
        $rows = foreach ($item in $aliases) {
            Register-Row $item
            $e = $item.Entry
            $label = $PowerShellActionLabels[[string]$e.action]
            if (-not $label) { $label = [string]$e.action }
            $paramText = (Get-ParameterChanges (Get-Field $e 'parameters')) -join '<br>'
            $hintText = & $useInstead $e
            if ($paramText) { $hintText = (@($paramText, $hintText) | Where-Object { $_ }) -join '<br>' }
            , @((Format-Code $e.old), (& $newCell $e), $label, $hintText)
        }
        Add-Table @('3.5 alias', '3.6 alias or command', 'Change', 'Use instead') $rows
        Add-DetailsEnd
    }
}

Add-Heading 1 'C# reference'
Add-Paragraph 'Changes per package, grouped by topic. Signatures are written as they appear in the 3.5 API; members not listed keep their names.'

foreach ($package in $packages) {
    if ($package.Changes.Count -eq 0) { continue }
    Add-Heading 2 $package.Name
    $categoryRank = { param($c) $i = [Array]::IndexOf([string[]]$CSharpCategoryOrder, [string]$c); if ($i -lt 0) { $i = 999 }; '{0:D3}' -f $i }
    $sorted = Sort-Ordinal $package.Changes { param($i) (& $categoryRank (Get-Field $i.Entry 'category')) + [char]1 + (Get-SortName ([string](Get-Field $i.Entry 'category'))) + [char]1 + (Get-SortName ([string]$i.Entry.old)) + [char]1 + $i.Id }
    $summary = "$(Get-Plural $package.Changes.Count 'change')"
    Add-DetailsStart $summary
    $rows = foreach ($item in $sorted) {
        Register-Row $item
        $e = $item.Entry
        $category = [string](Get-Field $e 'category')
        $categoryText = if ($CSharpCategoryTitles.ContainsKey($category)) { $CSharpCategoryTitles[$category] } elseif ($category) { Format-PlainText $category } else { '' }
        $label = $CSharpActionLabels[[string]$e.action]
        if (-not $label) { $label = [string]$e.action }
        $new = Get-Field $e 'new'
        $newText = if (Test-Text $new) { Format-Code $new } else { '&mdash;' }
        $detail = Get-Field $e 'replacement'
        if (-not (Test-Text $detail)) { $detail = Get-Field $e 'reason' }
        , @($categoryText, (Format-Code $e.old), $label, $newText, (Format-Prose ([string]$detail)))
    }
    Add-Table @('Topic', '3.5', 'Change', '3.6', 'Details') $rows
    Add-DetailsEnd
}

#endregion

#region Assertions

$problems = [System.Collections.Generic.List[string]]::new()
foreach ($item in @($allPsEntries) + @($allCsEntries)) {
    $count = if ($emitted.ContainsKey($item.Id)) { $emitted[$item.Id] } else { 0 }
    if ($count -ne 1) { $problems.Add("$($item.Id) ('$($item.Entry.old)') appears $count times in the reference tables.") }
}
$expected = $allPsEntries.Count + $allCsEntries.Count
$actual = ($emitted.Values | Measure-Object -Sum).Sum
if ($actual -ne $expected) { $problems.Add("Reference tables have $actual rows for $expected map entries.") }

# Every table row must have as many cells as its header.
$inFence = $null
$columns = -1
for ($i = 0; $i -lt $script:Lines.Count; $i++) {
    $line = $script:Lines[$i]
    $fenceMatch = [regex]::Match($line, '^(`{3,}|~{3,})')
    if ($fenceMatch.Success) {
        if ($null -eq $inFence) { $inFence = $fenceMatch.Value } elseif ($line.TrimEnd() -eq $inFence) { $inFence = $null }
        continue
    }
    if ($null -ne $inFence) { continue }
    if ($line.StartsWith('|')) {
        $cells = [regex]::Matches($line, '(?<!\\)\|').Count - 1
        if ($columns -lt 0) { $columns = $cells } elseif ($cells -ne $columns) { $problems.Add("Line $($i + 1): table row has $cells cells, header has $columns.") }
    } else {
        $columns = -1
    }
}
if ($null -ne $inFence) { $problems.Add('A code block is not closed.') }

if ($problems.Count -gt 0) {
    throw "Migration guide generation failed:`n" + ($problems -join "`n")
}

foreach ($warning in $script:DataWarnings) { Write-Warning $warning }
if ($FailOnDataWarning -and $script:DataWarnings.Count -gt 0) {
    throw "$($script:DataWarnings.Count) map data warnings."
}

#endregion

while ($script:Lines.Count -gt 0 -and $script:Lines[$script:Lines.Count - 1] -eq '') { $script:Lines.RemoveAt($script:Lines.Count - 1) }
$text = ($script:Lines -join "`n") + "`n"
$fullOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outputDirectory = Split-Path -Parent $fullOutput
if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory)) { [void](New-Item -ItemType Directory -Path $outputDirectory) }
[System.IO.File]::WriteAllText($fullOutput, $text, [System.Text.UTF8Encoding]::new($false))

[pscustomobject]@{
    Path              = $fullOutput
    Bytes             = [System.Text.Encoding]::UTF8.GetByteCount($text)
    Lines             = $script:Lines.Count
    PowerShellEntries = $allPsEntries.Count
    CSharpEntries     = $allCsEntries.Count
    DataWarnings      = $script:DataWarnings.Count
}
