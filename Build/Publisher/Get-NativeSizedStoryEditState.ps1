function Get-NativeSizedStoryRangeState {
    <#
    .SYNOPSIS
    Observes a complete uniform ordinary-story or table-cell range without editing it.
    .DESCRIPTION
    Whole-range font getters include final-CR typing and reject mixed values at
    the caller's explicit style boundary. Captures each producer's own paragraph,
    kerning and tab settings without imposing another producer's defaults.
    #>
    param($Range, [Collections.Generic.List[object]] $Objects)

    $font = $Range.Font
    $color = $font.Color
    $paragraph = $Range.ParagraphFormat
    $tabs = $paragraph.Tabs
    foreach ($wrapper in @($font, $color, $paragraph, $tabs)) {
        $Objects.Add($wrapper)
    }
    if ([int]$tabs.Count -ne 0) {
        throw 'Sized story editing requires the qualified tab-free paragraph profile.'
    }
    return [pscustomobject]@{
        text = [string]$Range.Text
        start = [int]$Range.Start
        end = [int]$Range.End
        length = [int]$Range.Length
        paragraphCount = [int]$Range.ParagraphsCount
        format = [pscustomobject]@{
            font = [string]$font.Name
            fontSizePoints = [double]$font.Size
            bold = [int]$font.Bold
            italic = [int]$font.Italic
            underline = [int]$font.Underline
            strikeThrough = [int]$font.StrikeThrough
            superScript = [int]$font.SuperScript
            subScript = [int]$font.SubScript
            rgb = [int]$color.RGB
            languageId = [int]$Range.LanguageID
            automaticPairKerningThresholdPoints = [double]$font.AutomaticPairKerningThreshold
            manualKerningPoints = [double]$font.Kerning
            alignment = [int]$paragraph.Alignment
            leftIndent = [double]$paragraph.LeftIndent
            rightIndent = [double]$paragraph.RightIndent
            firstLineIndent = [double]$paragraph.FirstLineIndent
            spaceBefore = [double]$paragraph.SpaceBefore
            spaceAfter = [double]$paragraph.SpaceAfter
            lineSpacingRule = [int]$paragraph.LineSpacingRule
            lineSpacingPoints = [double]$paragraph.LineSpacing
            listType = [int]$paragraph.ListType
            explicitTabCount = [int]$tabs.Count
        }
    }
}

function Assert-NativeSizedStoryRangeStyle {
    <#
    .SYNOPSIS
    Requires the sized recipe's ordinary uniform character and paragraph styling.
    #>
    param($State, [double] $Size, [double] $LineHeight, [bool] $Bold = $false)

    $format = $State.format
    $expectedBold = if ($Bold) { -1 } else { 0 }
    if ($State.paragraphCount -ne 1 -or $format.font -cne 'Arial' -or
        [Math]::Abs($format.fontSizePoints - $Size) -gt 0.001 -or
        $format.bold -ne $expectedBold -or $format.italic -ne 0 -or $format.underline -ne 0 -or
        $format.strikeThrough -ne 0 -or $format.superScript -ne 0 -or $format.subScript -ne 0 -or
        $format.rgb -ne 0 -or $format.languageId -ne 1033 -or $format.alignment -ne 0 -or
        $format.lineSpacingRule -ne 4 -or [Math]::Abs($format.lineSpacingPoints - $LineHeight) -gt 0.001 -or
        $format.leftIndent -ne 0 -or $format.rightIndent -ne 0 -or $format.firstLineIndent -ne 0 -or
        $format.spaceBefore -ne 0 -or $format.spaceAfter -ne 0 -or $format.listType -ne 255) {
        throw 'Sized story editing requires the recipe''s uniform ordinary text formatting, including its final CR.'
    }
}

function Get-NativeSizedStoryEditState {
    <#
    .SYNOPSIS
    Qualifies the two-frame sized composition and captures one-character edit evidence.
    .DESCRIPTION
    Admits only the three existing scalar size inputs and the exact N or M body.
    Resolves all ten semantic roles from actual owner, type, geometry and content,
    independent of producer-specific names. Preserves each input's own native
    names, IDs, reachable formatting and object observations. Complete story
    ranges are reacquired from both frames, with first/final-CR typing points;
    visible slices and the changed first character are recorded separately.
    Embedded bytes, native BStore sharing and unavailable story flags require
    the parent's saved-file comparisons. No native state or options are changed.
    .PARAMETER Publication
    The isolated worker's writable snapshot or read-only native reopen.
    .PARAMETER Objects
    Retains all acquired COM wrappers for the worker's lifetime.
    .PARAMETER ExpectedText
    Omit for the original N body; the only other accepted value is the complete
    M body with identical suffix and final CR.
    .EXAMPLE
    Get-NativeSizedStoryEditState -Publication $document -Objects $objects
    .OUTPUTS
    The current head Range for immediate editing, ShapeName and a serializable Snapshot.
    #>
    param($Publication, [Collections.Generic.List[object]] $Objects, [string] $ExpectedText)

    $setup = $Publication.PageSetup
    $pages = $Publication.Pages
    $masters = $Publication.MasterPages
    foreach ($wrapper in @($setup, $pages, $masters)) {
        $Objects.Add($wrapper)
    }
    $width = [double]$setup.PageWidth
    $height = [double]$setup.PageHeight
    $profiles = @('SizedCompositionA4', 'SizedCompositionPortrait', 'SizedCompositionLandscape') |
        ForEach-Object { Get-NativeSizedCompositionProfile -Case $_ }
    $matches = @($profiles | Where-Object {
        [Math]::Abs($_.widthPoints - $width) -le 0.001 -and
        [Math]::Abs($_.heightPoints - $height) -le 0.001
    })
    if ($matches.Count -ne 1 -or $pages.Count -ne 2 -or $masters.Count -ne 1) {
        throw 'Sized story editing requires exactly one qualified size, two physical pages and one centered master.'
    }
    $profile = $matches[0]
    $original = $profile.nativeLinkedText
    $edited = 'M' + $original.Substring(1)
    if ([string]::IsNullOrEmpty($ExpectedText)) {
        $ExpectedText = $original
    }
    if ($ExpectedText -cne $original -and $ExpectedText -cne $edited) {
        throw 'Sized story editing accepts only the exact 148-character N or M story, including its final CR.'
    }
    $master = $masters.Item(1)
    $Objects.Add($master)
    if ($master.IsTwoPageMaster) {
        throw 'Sized story editing requires a single centered master.'
    }
    $owners = [ordered]@{
        Master = $master
        Physical1 = $pages.Item(1)
        Physical2 = $pages.Item(2)
    }
    $ids = [Collections.Generic.HashSet[int]]::new()
    $ownerStates = [Collections.Generic.List[object]]::new()
    $flow = [Collections.Generic.List[object]]::new()
    $linkedFrames = @{}
    foreach ($ownerRole in $owners.Keys) {
        $owner = $owners[$ownerRole]
        $shapes = $owner.Shapes
        $Objects.Add($owner)
        $Objects.Add($shapes)
        $wantedShapes = @($profile.objects | Where-Object { $_.Owner -ceq $ownerRole })
        if ($shapes.Count -ne $wantedShapes.Count) {
            throw ('Sized story editing requires the complete 3/4/3 object ownership profile: ' + $ownerRole)
        }
        $assignment = $null
        if ($ownerRole -cne 'Master') {
            $assigned = $owner.Master
            $Objects.Add($assigned)
            if ([int]$assigned.PageID -ne [int]$master.PageID -or $owner.IgnoreMaster) {
                throw 'Sized story editing requires the centered master applied to both physical pages.'
            }
            $assignment = [int]$assigned.PageID
        }
        $shapeStates = [Collections.Generic.List[object]]::new()
        for ($index = 1; $index -le $shapes.Count; $index++) {
            $shape = $shapes.Item($index)
            $Objects.Add($shape)
            $id = [int]$shape.ID
            $wanted = $wantedShapes[$index - 1]
            if ($id -le 0 -or -not $ids.Add($id) -or [int]$shape.Type -ne $wanted.Type) {
                throw 'Sized story editing requires distinct physical shape IDs and the exact owner-specific role/type order.'
            }
            $bounds = @([double]$shape.Left, [double]$shape.Top, [double]$shape.Width, [double]$shape.Height)
            foreach ($coordinate in 0..3) {
                if (-not ([Math]::Abs($bounds[$coordinate] - $wanted.Bounds[$coordinate]) -le 0.001)) {
                    throw ('Sized story editing found incorrect role geometry: ' + $wanted.Name)
                }
            }
            $wrap = $shape.TextWrap
            $Objects.Add($wrap)
            $line = $null
            if ($wanted.Type -ne 18) {
                $line = $shape.Line
                $Objects.Add($line)
            }
            if ([int]$wrap.Type -ne 0 -or ($wanted.Type -ne 18 -and [int]$line.Visible -ne 0)) {
                throw ('Sized story editing requires None wrapping and no painted outline: ' + $wanted.Name)
            }
            $textState = $null
            $layout = $null
            $paint = $null
            $picture = $null
            $tableState = $null
            if ($wanted.Type -eq 17) {
                $fill = $shape.Fill
                $Objects.Add($fill)
                if ([int]$fill.Visible -ne 0) {
                    throw 'Sized story editing requires the existing unfilled ordinary text frames.'
                }
                $paint = [pscustomobject]@{ visible = [int]$fill.Visible }
                $frame = $shape.TextFrame
                $parent = $frame.Parent
                $story = $frame.Story
                $whole = $story.TextRange
                $visible = $frame.TextRange
                foreach ($wrapper in @($frame, $parent, $story, $whole, $visible)) {
                    $Objects.Add($wrapper)
                }
                if ([int]$parent.ID -ne $id) {
                    throw 'A sized text-frame parent differs from its enumerated physical shape ID.'
                }
                $linked = $wanted.Name -cin @('LinkedLead', 'LinkedContinuation')
                $wholeState = Get-NativeSizedStoryRangeState -Range $whole -Objects $Objects
                $expectedBody = if ($linked) { $ExpectedText } else { $wanted.Text + "`r" }
                if ($wholeState.text -cne $expectedBody -or $wholeState.length -ne $expectedBody.Length) {
                    throw ('A complete sized story differs from its qualified role text: ' + $wanted.Name)
                }
                Assert-NativeSizedStoryRangeStyle -State $wholeState -Size $wanted.Size -LineHeight $wanted.LineHeight -Bold $wanted.Bold
                $columns = $wanted.Name -ceq 'Page2Columns'
                $expectedColumns = if ($columns) { 2 } else { 1 }
                $expectedGap = if ($columns) { 6 } else { 0 }
                $expectedInset = if ($columns) { 3 } else { 0 }
                $hasNext = $frame.HasNextLink -eq -1
                $hasPrevious = $frame.HasPreviousLink -eq -1
                if ([int]$frame.Columns -ne $expectedColumns -or [int]$frame.VerticalTextAlignment -ne 0 -or
                    [Math]::Abs([double]$frame.ColumnSpacing - $expectedGap) -gt 0.001 -or
                    $hasNext -ne ($wanted.Name -ceq 'LinkedLead') -or
                    $hasPrevious -ne ($wanted.Name -ceq 'LinkedContinuation')) {
                    throw ('Sized story editing found incorrect columns, alignment or link flags: ' + $wanted.Name)
                }
                foreach ($property in @('MarginLeft', 'MarginTop', 'MarginRight', 'MarginBottom')) {
                    if ([Math]::Abs([double]$frame.$property - $expectedInset) -gt 0.001) {
                        throw ('Sized story editing found incorrect frame padding: ' + $wanted.Name)
                    }
                }
                $layout = [pscustomobject]@{
                    columns = [int]$frame.Columns
                    nativeColumnSpacing = [double]$frame.ColumnSpacing
                    left = [double]$frame.MarginLeft
                    top = [double]$frame.MarginTop
                    right = [double]$frame.MarginRight
                    bottom = [double]$frame.MarginBottom
                    verticalAlignment = [int]$frame.VerticalTextAlignment
                    hasNext = $hasNext
                    hasPrevious = $hasPrevious
                    parentShapeId = [int]$parent.ID
                }
                $textState = [pscustomobject]@{
                    text = if ($linked) { $wholeState.text.Substring(1) } else { $wholeState.text }
                    start = $wholeState.start
                    end = $wholeState.end
                    length = $wholeState.length
                    paragraphCount = $wholeState.paragraphCount
                    format = $wholeState.format
                }
                if ($linked) {
                    $typingPoints = [Collections.Generic.List[object]]::new()
                    $pointInvariants = [Collections.Generic.List[object]]::new()
                    foreach ($pointIndex in @(1, 148)) {
                        $character = $whole.Characters($pointIndex, 1)
                        $Objects.Add($character)
                        $pointState = Get-NativeSizedStoryRangeState -Range $character -Objects $Objects
                        $expectedCharacter = $ExpectedText.Substring($pointIndex - 1, 1)
                        if ($pointState.length -ne 1 -or $pointState.text -cne $expectedCharacter -or
                            ($pointState.format | ConvertTo-Json -Depth 4 -Compress) -cne
                            ($wholeState.format | ConvertTo-Json -Depth 4 -Compress)) {
                            throw 'Sized story editing requires observed first/final-CR typing with the complete body''s uniform format.'
                        }
                        $pointInvariants.Add([pscustomobject]@{
                            requestedIndex = $pointIndex
                            start = $pointState.start
                            end = $pointState.end
                            length = $pointState.length
                            format = $pointState.format
                        })
                        $typingPoints.Add([pscustomobject]@{
                            requestedIndex = $pointIndex
                            text = $pointState.text
                            length = $pointState.length
                            format = $pointState.format
                        })
                    }
                    $textState | Add-Member -NotePropertyName typingPoints -NotePropertyValue $pointInvariants.ToArray()
                    $linkedFrames[$wanted.Name] = [pscustomobject]@{
                        frame = $frame
                        range = $whole
                        shapeId = $id
                        shapeName = [string]$shape.Name
                    }
                    $flow.Add([pscustomobject]@{
                        role = $wanted.Name
                        shapeId = $id
                        wholeStoryText = $wholeState.text
                        visibleText = [string]$visible.Text
                        visibleStart = [int]$visible.Start
                        visibleEnd = [int]$visible.End
                        visibleLength = [int]$visible.Length
                        typingPoints = $typingPoints.ToArray()
                    })
                }
            } elseif ($wanted.Type -eq 1) {
                $fill = $shape.Fill
                $color = $fill.ForeColor
                $Objects.Add($fill)
                $Objects.Add($color)
                if ([int]$fill.Visible -ne -1 -or [int]$color.RGB -ne $wanted.Rgb) {
                    throw ('Sized story editing found incorrect rectangle paint: ' + $wanted.Name)
                }
                $paint = [pscustomobject]@{
                    visible = [int]$fill.Visible
                    type = [int]$fill.Type
                    rgb = [int]$color.RGB
                    transparency = [double]$fill.Transparency
                }
            } elseif ($wanted.Type -eq 13) {
                $format = $shape.PictureFormat
                $Objects.Add($format)
                $isEmpty = $format.IsEmpty
                $isLinked = $format.IsLinked
                $emptyType = if ($null -eq $isEmpty) { '<null>' } else { $isEmpty.GetType().FullName }
                $linkedType = if ($null -eq $isLinked) { '<null>' } else { $isLinked.GetType().FullName }
                $emptyValue = if ($null -eq $isEmpty) { '<null>' } else { [string]$isEmpty }
                $linkedValue = if ($null -eq $isLinked) { '<null>' } else { [string]$isLinked }
                if ($isEmpty -ne 0 -or $isLinked -ne 0) {
                    throw ('Sized story editing requires both existing embedded pictures. Role={0}; IsEmpty={1} ({2}); IsLinked={3} ({4}).' -f
                        $wanted.Name, $emptyValue, $emptyType, $linkedValue, $linkedType)
                }
                $picture = [pscustomobject]@{
                    fileName = [string]$format.FileName
                    isEmpty = [int]$isEmpty
                    isEmptyRawValue = $emptyValue
                    isEmptyRuntimeType = $emptyType
                    isLinked = [int]$isLinked
                    isLinkedRawValue = $linkedValue
                    isLinkedRuntimeType = $linkedType
                    cropLeft = [double]$format.CropLeft
                    cropTop = [double]$format.CropTop
                    cropRight = [double]$format.CropRight
                    cropBottom = [double]$format.CropBottom
                }
                if ($picture.cropLeft -ne 0 -or $picture.cropTop -ne 0 -or $picture.cropRight -ne 0 -or $picture.cropBottom -ne 0) {
                    throw 'Sized story editing requires the qualified uncropped picture placements.'
                }
            } else {
                $geometry = Get-NativeTableGeometry $shape
                $cells = @(Get-NativeTableCellInventory -Shape $shape -Objects $Objects)
                if ($geometry.columnWidthsPoints.Count -ne 2 -or $geometry.rowHeightsPoints.Count -ne 2 -or $cells.Count -ne 4) {
                    throw 'Sized story editing requires the original two-by-two calendar.'
                }
                foreach ($track in 0..1) {
                    if ([Math]::Abs($geometry.columnWidthsPoints[$track] - $profile.columnWidths[$track]) -gt 0.001 -or
                        [Math]::Abs($geometry.rowHeightsPoints[$track] - $profile.rowHeights[$track]) -gt 0.001) {
                        throw 'Sized story editing found incorrect calendar tracks.'
                    }
                }
                $table = $shape.Table
                $Objects.Add($table)
                for ($cellIndex = 0; $cellIndex -lt 4; $cellIndex++) {
                    $row = if ($cellIndex -lt 2) { 1 } else { 2 }
                    $column = ($cellIndex % 2) + 1
                    $cellRange = $table.Cells($row, $column, $row, $column)
                    $cell = $cellRange.Item(1)
                    $range = $cell.TextRange
                    foreach ($wrapper in @($cellRange, $cell, $range)) {
                        $Objects.Add($wrapper)
                    }
                    $rangeState = Get-NativeSizedStoryRangeState -Range $range -Objects $Objects
                    Assert-NativeSizedStoryRangeStyle -State $rangeState -Size 9 -LineHeight 12 -Bold ($row -eq 1)
                    if ($cells[$cellIndex].text -cne ($profile.cellText[$cellIndex] + "`r") -or
                        $rangeState.text -cne ($profile.cellText[$cellIndex] + "`r") -or
                        $cells[$cellIndex].row -ne $row -or $cells[$cellIndex].column -ne $column -or
                        $cells[$cellIndex].rowSpan -ne 1 -or $cells[$cellIndex].columnSpan -ne 1 -or
                        $cells[$cellIndex].style.fill.status -cne 'Observed' -or $cells[$cellIndex].style.fill.visible -ne 0) {
                        throw 'Sized story editing found changed calendar text, spans or fill state.'
                    }
                    # The existing observer retains an unassessed diagonal;
                    # this borderless recipe qualifies only the four outer edges.
                    foreach ($side in @('left', 'top', 'right', 'bottom')) {
                        $border = $cells[$cellIndex].style.borders[$side]
                        if ($border.status -cne 'Observed' -or $border.widthPoints -ne 0) {
                            throw 'Sized story editing requires four observed zero-width calendar borders.'
                        }
                    }
                    $cells[$cellIndex] | Add-Member -NotePropertyName rangeState -NotePropertyValue $rangeState
                    $padding = [pscustomobject]@{
                        left = [double]$cell.MarginLeft
                        top = [double]$cell.MarginTop
                        right = [double]$cell.MarginRight
                        bottom = [double]$cell.MarginBottom
                    }
                    if ($padding.left -ne 0 -or $padding.top -ne 0 -or $padding.right -ne 0 -or $padding.bottom -ne 0) {
                        throw 'Sized story editing requires the qualified zero calendar-cell margins.'
                    }
                    $cells[$cellIndex] | Add-Member -NotePropertyName padding -NotePropertyValue $padding
                }
                $tableState = [pscustomobject]@{
                    geometry = $geometry
                    cells = $cells
                }
            }
            $shapeStates.Add([pscustomobject]@{
                role = $wanted.Name
                nativeShapeId = $id
                actualName = [string]$shape.Name
                type = [int]$shape.Type
                bounds = $bounds
                wrapping = [int]$wrap.Type
                outlineVisible = if ($wanted.Type -eq 18) { $null } else { [int]$line.Visible }
                text = $textState
                layout = $layout
                paint = $paint
                picture = $picture
                table = $tableState
            })
        }
        $ownerStates.Add([pscustomobject]@{
            role = $ownerRole
            pageId = [int]$owner.PageID
            actualName = [string]$owner.Name
            pageType = [int]$owner.PageType
            masterPageId = $assignment
            shapes = $shapeStates.ToArray()
        })
    }
    if ($ids.Count -ne 10 -or $linkedFrames.Count -ne 2) {
        throw 'Sized story editing did not cover exactly ten objects and its two linked frames.'
    }
    $next = $linkedFrames.LinkedLead.frame.NextLinkedTextFrame
    $previous = $linkedFrames.LinkedContinuation.frame.PreviousLinkedTextFrame
    $nextParent = $next.Parent
    $previousParent = $previous.Parent
    foreach ($wrapper in @($next, $previous, $nextParent, $previousParent)) {
        $Objects.Add($wrapper)
    }
    if ([int]$nextParent.ID -ne $linkedFrames.LinkedContinuation.shapeId -or
        [int]$previousParent.ID -ne $linkedFrames.LinkedLead.shapeId) {
        throw 'Sized story editing requires reciprocal links covering exactly its enumerated page 1/2 chain.'
    }
    return [pscustomobject]@{
        Range = $linkedFrames.LinkedLead.range
        ShapeName = $linkedFrames.LinkedLead.shapeName
        Snapshot = [pscustomobject]@{
            schema = 'officeimo.publisher.native-sized-story-edit-state.v1'
            text = $ExpectedText
            invariant = [pscustomobject]@{
                profile = $profile.case
                requestedWidthEmu = $profile.widthEmu
                requestedHeightEmu = $profile.heightEmu
                nativeWidthPoints = $width
                nativeHeightPoints = $height
                defaultTabStopPoints = [double]$Publication.DefaultTabStop
                masterIsTwoPage = [bool]$master.IsTwoPageMaster
                uniqueNativeShapeCount = $ids.Count
                chainShapeIds = @($linkedFrames.LinkedLead.shapeId, $linkedFrames.LinkedContinuation.shapeId)
                owners = $ownerStates.ToArray()
            }
            flow = $flow.ToArray()
            savedImageBytesAndStoryFlags = 'Require separate saved-file reader comparison; no native image/catalog or hyphenation mutation is performed.'
        }
    }
}

function Assert-NativeSizedStoryEditState {
    <#
    .SYNOPSIS
    Requires only the admitted N-to-M text change across sized-story snapshots.
    .DESCRIPTION
    Compares all captured ownership, identities, names, geometry, formatting and
    unaffected content within the same input. Visible slices remain separate
    because the changed glyph can alter native flow. Both complete frame story
    observations must still contain the expected 148-character edited text and
    their first/final-CR typing points with invariant formatting.
    .PARAMETER Before
    Serializable state captured immediately before selecting the first N.
    .PARAMETER After
    Serializable state captured after editing or native save/reopen.
    .EXAMPLE
    Assert-NativeSizedStoryEditState -Before $beforeState.Snapshot -After $afterState.Snapshot
    .OUTPUTS
    None. Throws if captured state changes beyond the first character's text.
    #>
    param($Before, $After)

    $schema = 'officeimo.publisher.native-sized-story-edit-state.v1'
    if ($Before.schema -cne $schema -or $After.schema -cne $schema -or
        $Before.text.Length -ne 148 -or $Before.text[0] -cne 'N' -or
        $After.text -cne ('M' + $Before.text.Substring(1)) -or
        $Before.text[147] -cne "`r" -or $Before.flow.Count -ne 2 -or $After.flow.Count -ne 2) {
        throw 'Sized story edit snapshots must describe only the qualified complete N-to-M change.'
    }
    if (($Before.invariant | ConvertTo-Json -Depth 20 -Compress) -cne
        ($After.invariant | ConvertTo-Json -Depth 20 -Compress)) {
        throw 'Sized story editing changed captured ownership, native identities, formatting, layout or unaffected content.'
    }
    foreach ($state in @($Before, $After)) {
        foreach ($frame in $state.flow) {
            if ($frame.wholeStoryText -cne $state.text -or $frame.typingPoints.Count -ne 2) {
                throw 'Every sized linked frame must resolve to the same complete expected story.'
            }
            $role = @($state.invariant.owners.shapes | Where-Object { $_.role -ceq $frame.role })
            if ($role.Count -ne 1) {
                throw 'Sized typing-point observations must identify one preserved frame role.'
            }
            foreach ($pointOffset in 0..1) {
                $point = $frame.typingPoints[$pointOffset]
                $expectedIndex = if ($pointOffset -eq 0) { 1 } else { 148 }
                if ($point.requestedIndex -ne $expectedIndex -or $point.length -ne 1 -or
                    $point.text -cne $state.text.Substring($expectedIndex - 1, 1) -or
                    ($point.format | ConvertTo-Json -Depth 4 -Compress) -cne
                    ($role[0].text.format | ConvertTo-Json -Depth 4 -Compress)) {
                    throw 'Sized first/final-CR typing must retain its exact character and uniform body formatting.'
                }
            }
        }
    }
}
