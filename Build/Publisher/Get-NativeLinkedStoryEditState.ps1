function Get-NativeLinkedStoryEditState {
    <#
    .SYNOPSIS
    Validates the qualified three-frame ordinary story and captures edit evidence.
    .DESCRIPTION
    Accepts writable or read-only snapshots. Enumerated physical shape IDs must
    exactly cover the bounded page 1/3/2 chain; returned frame parents are checked
    against that inventory. Story ranges and typing observations are reacquired.
    The returned Range is for the caller's immediate edit only; Snapshot contains
    plain data suitable for the operation report. No native state is changed.
    .PARAMETER Publication
    The isolated worker's already opened candidate or native resave.
    .PARAMETER Objects
    Retains COM wrappers for the worker's lifetime.
    .PARAMETER ExpectedText
    The complete expected story after an edit. Omit for the original M heading.
    .OUTPUTS
    The head's current native range, shape name and serializable state snapshot.
    #>
    param($Publication, [Collections.Generic.List[object]] $Objects, [string] $ExpectedText)

    $canonicalText = 'Multi-paragraph heading. Body follows below.' + "`r`r" + 'Final paragraph: ' +
        ('Complete ordinary stories retain every paragraph and all overflow text. ' * 18).TrimEnd() + "`r"
    if ([string]::IsNullOrEmpty($ExpectedText)) {
        $ExpectedText = $canonicalText
    }
    if ($ExpectedText.Length -ne 1359 -or $ExpectedText.Substring(1) -cne $canonicalText.Substring(1)) {
        throw 'Linked story editing requires the qualified three-paragraph text profile.'
    }
    $setup = $Publication.PageSetup
    $Objects.Add($setup)
    $pages = $Publication.Pages
    $Objects.Add($pages)
    if ($pages.Count -ne 3 -or $setup.PageWidth -ne 612 -or $setup.PageHeight -ne 792) {
        throw 'Linked story editing requires three Letter-size physical pages.'
    }
    $frames = @{}
    $snapshots = [Collections.Generic.List[object]]::new()
    $head = $null
    for ($pageIndex = 1; $pageIndex -le 3; $pageIndex++) {
        $page = $pages.Item($pageIndex)
        $Objects.Add($page)
        $shapes = $page.Shapes
        $Objects.Add($shapes)
        if ($shapes.Count -ne 1) {
            throw 'Linked story editing requires exactly one physical text box per page.'
        }
        $shape = $shapes.Item(1)
        $Objects.Add($shape)
        if ($shape.Type -ne 17 -or $shape.HasTextFrame -ne -1) {
            throw 'Linked story editing accepts only the ordinary three-text-box profile.'
        }
        $top = if ($pageIndex -eq 1) { 72 } elseif ($pageIndex -eq 3) { 144 } else { 216 }
        $width = if ($pageIndex -eq 2) { 360 } else { 216 }
        $height = if ($pageIndex -eq 2) { 108 } else { 72 }
        if ($shape.Left -ne 72 -or $shape.Top -ne $top -or $shape.Width -ne $width -or $shape.Height -ne $height) {
            throw 'Linked story frame geometry differs from the qualified page 1/3/2 profile.'
        }
        $id = [int]$shape.ID
        if ($id -le 0 -or $frames.ContainsKey($id)) {
            throw 'Linked story physical shape IDs must be unique.'
        }
        $frame = $shape.TextFrame
        $Objects.Add($frame)
        $parent = $frame.Parent
        $Objects.Add($parent)
        if ([int]$parent.ID -ne $id) {
            throw 'A text-frame parent does not resolve to its enumerated physical shape ID.'
        }
        $hasPrevious = $frame.HasPreviousLink -eq -1
        $hasNext = $frame.HasNextLink -eq -1
        if ($hasPrevious -ne ($pageIndex -ne 1) -or $hasNext -ne ($pageIndex -ne 2)) {
            throw 'Linked story head, middle and tail flags differ from the qualified chain.'
        }
        $story = $frame.Story
        $Objects.Add($story)
        $range = $story.TextRange
        $Objects.Add($range)
        if ([string]$range.Text -cne $ExpectedText -or [int]$range.Length -ne 1359 -or $range.ParagraphsCount -ne 3) {
            throw 'Every physical frame must resolve to the same complete three-paragraph story.'
        }
        $frames[$id] = [pscustomobject]@{ shape = $shape; frame = $frame; range = $range; pageIndex = $pageIndex }
        if ($pageIndex -eq 1) {
            $head = $frames[$id]
        }
        $snapshots.Add([pscustomobject]@{
            pageIndex = $pageIndex
            pageId = [int]$page.PageID
            shapeId = $id
            left = [double]$shape.Left
            top = [double]$shape.Top
            width = [double]$shape.Width
            height = [double]$shape.Height
            hasPrevious = $hasPrevious
            hasNext = $hasNext
            columns = [int]$frame.Columns
            columnSpacing = [double]$frame.ColumnSpacing
            marginLeft = [double]$frame.MarginLeft
            marginTop = [double]$frame.MarginTop
            marginRight = [double]$frame.MarginRight
            marginBottom = [double]$frame.MarginBottom
            verticalAlignment = [int]$frame.VerticalTextAlignment
        })
    }
    $visited = [Collections.Generic.HashSet[int]]::new()
    $chain = [Collections.Generic.List[int]]::new()
    $current = $head
    $previousId = $null
    foreach ($expectedPage in @(1, 3, 2)) {
        $id = [int]$current.shape.ID
        if (-not $visited.Add($id) -or $current.pageIndex -ne $expectedPage) {
            throw 'The native linked story has a repeated or incorrectly ordered physical frame.'
        }
        $chain.Add($id)
        if ($null -ne $previousId) {
            $previous = $current.frame.PreviousLinkedTextFrame
            $Objects.Add($previous)
            $previousParent = $previous.Parent
            $Objects.Add($previousParent)
            if ([int]$previousParent.ID -ne $previousId) {
                throw 'A linked frame backward reference does not identify its preceding physical shape.'
            }
        }
        if ($expectedPage -ne 2) {
            $next = $current.frame.NextLinkedTextFrame
            $Objects.Add($next)
            $nextParent = $next.Parent
            $Objects.Add($nextParent)
            $nextId = [int]$nextParent.ID
            if (-not $frames.ContainsKey($nextId)) {
                throw 'A linked frame points outside the enumerated physical shape inventory.'
            }
            $previousId = $id
            $current = $frames[$nextId]
        }
    }
    if ($visited.Count -ne $frames.Count) {
        throw 'The reachable linked story does not cover every physical frame.'
    }
    $observation = Get-NativeMultiParagraphObservation -Shape $head.shape -ExpectedText $ExpectedText -Stage 'LinkedStoryEditState'
    if (-not $observation.storyRange.matchesExpectedText -or
        ($observation.storyRange.crIndices -join ',') -ne '45,46,1359' -or
        $observation.points.Count -ne 11 -or $observation.paragraphs.Count -ne 3 -or
        @($observation.points | Where-Object status -ne 'Observed').Count -ne 0 -or
        @($observation.paragraphs | Where-Object status -ne 'Observed').Count -ne 0) {
        throw 'Linked story typing points or paragraph observations are incomplete.'
    }
    return [pscustomobject]@{
        Range = $head.range
        ShapeName = [string]$head.shape.Name
        Snapshot = [pscustomobject]@{
            text = $ExpectedText
            frames = $snapshots.ToArray()
            chainShapeIds = $chain.ToArray()
            observation = $observation
        }
    }
}

function Assert-NativeLinkedStoryEditState {
    <#
    .SYNOPSIS
    Requires unchanged chain identity, point styling and paragraph formatting.
    .PARAMETER Before
    Serializable state captured before the one-character edit.
    .PARAMETER After
    Serializable state captured immediately after editing or native reopen.
    .OUTPUTS
    None. Refuses differences other than the requested first character's text.
    #>
    param($Before, $After)

    if (($Before.frames | ConvertTo-Json -Depth 4 -Compress) -cne ($After.frames | ConvertTo-Json -Depth 4 -Compress) -or
        ($Before.chainShapeIds -join ',') -ne ($After.chainShapeIds -join ',')) {
        throw 'Native linked story editing changed physical frame identity, geometry or layout.'
    }
    for ($index = 0; $index -lt 11; $index++) {
        $beforePoint = $Before.observation.points[$index]
        $afterPoint = $After.observation.points[$index]
        foreach ($name in @('role', 'requestedIndex', 'status', 'start', 'end', 'length', 'font', 'fontSizePoints', 'bold', 'rgb', 'languageId')) {
            if ($beforePoint.$name -cne $afterPoint.$name) {
                throw ('Native linked story editing changed a typing-point property: ' + $name)
            }
        }
    }
    for ($index = 0; $index -lt 3; $index++) {
        $beforeParagraph = $Before.observation.paragraphs[$index]
        $afterParagraph = $After.observation.paragraphs[$index]
        foreach ($name in @('ordinal', 'status', 'start', 'end', 'length', 'font', 'fontSizePoints', 'bold', 'languageId',
            'alignment', 'leftIndent', 'rightIndent', 'firstLineIndent', 'spaceBefore', 'spaceAfter', 'lineSpacingRule', 'lineSpacingPoints', 'listType')) {
            if ($beforeParagraph.$name -cne $afterParagraph.$name) {
                throw ('Native linked story editing changed a paragraph property: ' + $name)
            }
        }
    }
}
