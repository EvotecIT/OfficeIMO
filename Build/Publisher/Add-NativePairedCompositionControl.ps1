function Set-NativePairedCompositionTextStyle {
    <#
    .SYNOPSIS
    Applies the paired composition's explicit ordinary-run and paragraph style.
    .DESCRIPTION
    Accepts a freshly acquired complete story or cell range, including its final
    CR. Uses the table-creation control's fresh-range formatting sequence with
    this composition's 16-point line height and English language. Retains COM
    wrappers in the isolated worker; no application or story options are changed.
    #>
    param($Range, [bool] $Bold = $false)

    $font = $Range.Font
    $color = $font.Color
    $paragraph = $Range.ParagraphFormat
    $objects.Add($font)
    $objects.Add($color)
    $objects.Add($paragraph)
    $font.Name = 'Arial'
    $font.Size = 12
    $font.Bold = if ($Bold) { -1 } else { 0 }
    $font.Italic = 0
    $color.RGB = 0
    $Range.LanguageID = 1033
    $paragraph.Alignment = 0
    $paragraph.LeftIndent = 0
    $paragraph.RightIndent = 0
    $paragraph.FirstLineIndent = 0
    $paragraph.SpaceBefore = 0
    $paragraph.SpaceAfter = 0
    $paragraph.SetLineSpacing(4, 16)
}

function Add-NativePairedCompositionTextFrame {
    <#
    .SYNOPSIS
    Adds one plain, unwrapped paired-composition frame using the worker's authoring helper.
    .DESCRIPTION
    The complete story range is reacquired after text assignment so formatting
    includes the final CR. Column and inset changes remain with the caller.
    #>
    param($Page, [string] $Name, [string] $Text, [double[]] $Bounds)

    $shape = Add-TextFrame $Page $Name $Text $Bounds[0] $Bounds[1] $Bounds[2] $Bounds[3] 12
    $frame = $shape.TextFrame
    $story = $frame.Story
    $range = $story.TextRange
    $wrap = $shape.TextWrap
    $objects.Add($frame)
    $objects.Add($story)
    $objects.Add($range)
    $objects.Add($wrap)
    $frame.Columns = 1
    $frame.ColumnSpacing = 0
    $frame.VerticalTextAlignment = 0
    $wrap.Type = 0
    Set-NativePairedCompositionTextStyle -Range $range
    return $shape
}

function Assert-NativePairedCompositionMasterOwnership {
    <#
    .SYNOPSIS
    Verifies the three authored shapes in each actual paired-composition master.
    .DESCRIPTION
    Enumerates each native master's Shapes collection independently of the
    authoring return values and expected-shape rows. Requires the intended
    marker, rule and logo names, distinct native shape IDs and explicit bounds.
    This is a read-only guard for the PairedBrochure control, not a file-codec
    assessment. Retains acquired COM wrappers in the isolated worker's list.
    .PARAMETER Document
    The task-owned PairedBrochure publication after authoring or native save.
    .EXAMPLE
    Assert-NativePairedCompositionMasterOwnership -Document $document
    .OUTPUTS
    Serializable master and native shape identities, names, types and observed bounds.
    #>
    param($Document)

    $masters = $Document.MasterPages
    $objects.Add($masters)
    if ($masters.Count -ne 3) {
        throw 'PairedBrochure requires exactly three actual native master definitions.'
    }
    $markerNames = @('DefaultMasterMarker', 'LeftMasterMarker', 'RightMasterMarker')
    $shapeIds = [Collections.Generic.HashSet[int]]::new()
    $masterObservations = [Collections.Generic.List[object]]::new()
    for ($masterIndex = 1; $masterIndex -le 3; $masterIndex++) {
        $master = $masters.Item($masterIndex)
        $shapes = $master.Shapes
        $objects.Add($master)
        $objects.Add($shapes)
        $expectedPair = $masterIndex -ne 1
        if ([bool]$master.IsTwoPageMaster -ne $expectedPair -or $shapes.Count -ne 3) {
            throw ('PairedBrochure native master ' + $masterIndex + ' must own exactly its marker, rule and logo.')
        }
        $expectedNames = @(
            $markerNames[$masterIndex - 1],
            ('PairedBrochureMasterRule' + $masterIndex),
            ('PairedBrochureMasterLogo' + $masterIndex)
        )
        $seenNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $shapeObservations = [Collections.Generic.List[object]]::new()
        for ($shapeIndex = 1; $shapeIndex -le $shapes.Count; $shapeIndex++) {
            $shape = $shapes.Item($shapeIndex)
            $objects.Add($shape)
            $name = [string]$shape.Name
            $shapeId = [int]$shape.ID
            if ($expectedNames -cnotcontains $name -or -not $seenNames.Add($name)) {
                throw ('Unexpected or duplicate shape on PairedBrochure native master ' + $masterIndex + ': ' + $name)
            }
            if ($shapeId -le 0 -or -not $shapeIds.Add($shapeId)) {
                throw ('PairedBrochure native shape IDs must be positive and unique across masters: ' + $shapeId)
            }
            if ($name -ceq $expectedNames[0]) {
                $expectedType = 17
                $markerLeft = if ($masterIndex -eq 3) { 342 } else { 18 }
                $expectedBounds = @($markerLeft, 24, 252, 24)
            } elseif ($name -ceq $expectedNames[1]) {
                $expectedType = 1
                $expectedBounds = @(36, 66, 540, 6)
            } else {
                $expectedType = 13
                $logoLeft = if ($masterIndex -eq 2) { 72 } else { 504 }
                $expectedBounds = @($logoLeft, 708, 72, 48)
            }
            if ([int]$shape.Type -ne $expectedType) {
                throw ('Unexpected native shape type on PairedBrochure master ' + $masterIndex + ': ' + $name)
            }
            $properties = @('Left', 'Top', 'Width', 'Height')
            for ($coordinate = 0; $coordinate -lt $properties.Count; $coordinate++) {
                $property = $properties[$coordinate]
                $actual = [double]$shape.$property
                if ([Math]::Abs($actual - $expectedBounds[$coordinate]) -gt 0.001) {
                    throw ('Unexpected ' + $property + ' on PairedBrochure native master ' + $masterIndex + ': ' + $name)
                }
            }
            $shapeObservations.Add([pscustomobject]@{
                nativeShapeId = $shapeId
                name = $name
                type = [int]$shape.Type
                left = [double]$shape.Left
                top = [double]$shape.Top
                width = [double]$shape.Width
                height = [double]$shape.Height
            })
        }
        $masterObservations.Add([pscustomobject]@{
            masterIndex = $masterIndex
            masterPageId = [int]$master.PageID
            masterName = [string]$master.Name
            isTwoPageMaster = [bool]$master.IsTwoPageMaster
            shapeCount = [int]$shapes.Count
            shapes = $shapeObservations.ToArray()
        })
    }
    return [pscustomobject]@{
        schema = 'officeimo.publisher.native-paired-master-ownership.v1'
        status = 'Observed'
        masterCount = $masterObservations.Count
        uniqueNativeShapeCount = $shapeIds.Count
        coordinateTolerancePoints = 0.001
        masters = $masterObservations.ToArray()
    }
}

function Add-NativePairedCompositionControl {
    <#
    .SYNOPSIS
    Adds a mixed-content brochure to the independent native paired-master control.
    .DESCRIPTION
    Calls the existing TwoPageMaster authoring owner in a fresh publication and
    retains its centered-page geometry, unchanged markers and assignments. Adds
    master artwork, a physical-page 1/3/2 story chain, a borderless calendar and
    ordinary frame layouts. One synthetic PNG is embedded at four placements;
    saved native image-store sharing and story hyphenation require separate
    decoding. Application options are observed without writes.
    .PARAMETER Document
    The isolated worker's newly created, task-owned Publisher publication.
    .PARAMETER FirstPage
    Its first physical page before the base control adds pages two and three.
    .PARAMETER PicturePath
    Unused task-owned path for the synthetic 32-pixel reference PNG.
    .EXAMPLE
    Add-NativePairedCompositionControl -Document $document -FirstPage $page -PicturePath $picturePath
    .OUTPUTS
    Base page/master assignments, the composition contract and complete lead-story observations.
    #>
    param($Document, $FirstPage, [string] $PicturePath)

    $layout = Add-NativeComparisonLayout -Case TwoPageMaster -Document $Document -FirstPage $FirstPage
    $pages = $Document.Pages
    $masters = $Document.MasterPages
    $objects.Add($pages)
    $objects.Add($masters)
    if ($pages.Count -ne 3 -or $masters.Count -ne 3) {
        throw 'The paired composition requires the three-page, three-master native base control.'
    }
    $second = $pages.Item(2)
    $third = $pages.Item(3)
    $objects.Add($second)
    $objects.Add($third)

    Set-NativePhase 'PairedBrochure-CreateReferencePicture'
    New-ReferencePicture $PicturePath
    $colors = @(16711680, 255, 32768) # Native RGB integers: blue, red, green (0,128,0).
    for ($index = 1; $index -le $masters.Count; $index++) {
        Set-NativePhase ('PairedBrochure-MasterArtwork-' + $index)
        $master = $masters.Item($index)
        $objects.Add($master)
        $rule = $master.Shapes.AddShape(1, 36, 66, 540, 6)
        $rule.Name = 'PairedBrochureMasterRule' + $index
        $rule.Line.Visible = 0
        $rule.Fill.ForeColor.RGB = $colors[$index - 1]
        # Keep the left member's picture inside its half of the native spread;
        # placing it at 504 transfers ownership to the right member.
        $logoLeft = if ($index -eq 2) { 72 } else { 504 }
        $logo = $master.Shapes.AddPicture($PicturePath, 0, -1, $logoLeft, 708, 72, 48)
        $logo.Name = 'PairedBrochureMasterLogo' + $index
        $logo.Line.Visible = 0
        $objects.Add($rule)
        $objects.Add($logo)
        foreach ($shape in @($rule, $logo)) {
            $wrap = $shape.TextWrap
            $objects.Add($wrap)
            $wrap.Type = 0
        }
        $expectedShapes.Add([pscustomobject]@{
            name = $rule.Name
            kind = 'Rectangle'
            pageRole = 'Master'
            pageIndex = $index
            left = 36
            top = 66
            width = 540
            height = 6
            rgb = $colors[$index - 1]
        })
        $expectedShapes.Add([pscustomobject]@{
            name = $logo.Name
            kind = 'Image'
            pageRole = 'Master'
            pageIndex = $index
            left = $logoLeft
            top = 708
            width = 72
            height = 48
        })
    }

    Set-NativePhase 'PairedBrochure-AuthorLeadStory'
    $text = 'Multi-paragraph heading. Body follows below.' + "`r`r" + 'Final paragraph: ' +
        ('Complete ordinary stories retain every paragraph and all overflow text. ' * 18).TrimEnd() + "`r"
    if ($text.Length -ne 1359 -or $text[44] -ne "`r" -or $text[45] -ne "`r" -or $text[1358] -ne "`r") {
        throw 'The paired brochure lead story must retain the 45/46/1359 CR endpoints.'
    }
    $profiles = @(
        @{
            Page = $FirstPage
            PageIndex = 1
            Name = 'PairedBrochureLead'
            Bounds = @(36, 180, 330, 180)
        },
        @{
            Page = $third
            PageIndex = 3
            Name = 'PairedBrochureMiddle'
            Bounds = @(36, 180, 540, 180)
        },
        @{
            Page = $second
            PageIndex = 2
            Name = 'PairedBrochureEnd'
            Bounds = @(36, 180, 330, 180)
        }
    )
    $frames = [Collections.Generic.List[object]]::new()
    $lead = $null
    foreach ($profile in $profiles) {
        $inputText = if ($frames.Count -eq 0) { $text.Substring(0, $text.Length - 1) } else { '' }
        $shape = Add-NativePairedCompositionTextFrame -Page $profile.Page -Name $profile.Name -Text $inputText -Bounds $profile.Bounds
        if ($null -eq $lead) {
            $lead = $shape
        }
        $frame = $shape.TextFrame
        $objects.Add($frame)
        $frames.Add($frame)
        $expected = $expectedShapes[$expectedShapes.Count - 1]
        $expected.size = $null
        $expected | Add-Member -NotePropertyName pageIndex -NotePropertyValue $profile.PageIndex
    }
    $story = $frames[0].Story
    $whole = $story.TextRange
    $objects.Add($story)
    $objects.Add($whole)
    if ([string]$whole.Text -cne $text -or [int]$whole.Length -ne $text.Length) {
        throw 'The reacquired paired brochure story does not contain the complete text and final CR.'
    }
    foreach ($run in @(
        @{
            Start = 1
            Length = 25
            Size = 16
        },
        @{
            Start = 46
            Length = 1
            Size = 14
        },
        @{
            Start = 47
            Length = 17
            Size = 12
        }
    )) {
        $range = $whole.Characters($run.Start, $run.Length)
        $font = $range.Font
        $objects.Add($range)
        $objects.Add($font)
        $font.Size = $run.Size
        $font.Bold = -1
    }
    $lineHeights = @(22, 20, 18)
    foreach ($ordinal in 1..3) {
        $range = $whole.Paragraphs($ordinal, 1)
        $format = $range.ParagraphFormat
        $objects.Add($range)
        $objects.Add($format)
        $format.SetLineSpacing(4, $lineHeights[$ordinal - 1])
        if ([int]$format.ListType -ne 255) {
            throw 'The paired brochure requires ordinary lead paragraphs without list labels.'
        }
    }
    $observations = [Collections.Generic.List[object]]::new()
    $observations.Add((Get-NativeMultiParagraphObservation -Shape $lead -ExpectedText $text -Stage 'BeforeLinking'))
    $frames[0].NextLinkedTextFrame = $frames[1]
    $frames[1].NextLinkedTextFrame = $frames[2]
    if ($frames[0].HasNextLink -ne -1 -or $frames[0].HasPreviousLink -ne 0 -or
        $frames[1].HasNextLink -ne -1 -or $frames[1].HasPreviousLink -ne -1 -or
        $frames[2].HasNextLink -ne 0 -or $frames[2].HasPreviousLink -ne -1) {
        throw 'Publisher did not retain the paired brochure physical-page 1/3/2 chain.'
    }
    $observations.Add((Get-NativeMultiParagraphObservation -Shape $lead -ExpectedText $text -Stage 'AfterLinking'))

    Set-NativePhase 'PairedBrochure-AuthorFirstPage'
    $null = Add-NativePairedCompositionTextFrame -Page $FirstPage -Name 'PairedBrochureSidebar' -Text 'In this brochure: shared spaces, community projects and the events calendar.' -Bounds @(390, 180, 186, 156)
    $rule = $FirstPage.Shapes.AddShape(1, 36, 384, 540, 6)
    $rule.Name = 'PairedBrochurePhysicalRule'
    $rule.Line.Visible = 0
    $rule.Fill.ForeColor.RGB = 16711680
    $objects.Add($rule)
    $wrap = $rule.TextWrap
    $objects.Add($wrap)
    $wrap.Type = 0
    $expectedShapes.Add([pscustomobject]@{
        name = $rule.Name
        kind = 'Rectangle'
        pageIndex = 1
        left = 36
        top = 384
        width = 540
        height = 6
        rgb = 16711680
    })

    Set-NativePhase 'PairedBrochure-AuthorCalendar'
    $calendar = $FirstPage.Shapes.AddTable(2, 2, 36, 420, 330, 72)
    $calendar.Name = 'PairedBrochureCalendar'
    $table = $calendar.Table
    $columns = $table.Columns
    $rows = $table.Rows
    $objects.Add($calendar)
    $objects.Add($table)
    $objects.Add($columns)
    $objects.Add($rows)
    $table.ApplyAutoFormat(-2, $false, $false, $true, $true) # Existing borderless table-creation control.
    foreach ($index in 1..2) {
        $column = $columns.Item($index)
        $row = $rows.Item($index)
        $objects.Add($column)
        $objects.Add($row)
        $column.Width = @(180, 150)[$index - 1]
        $row.Height = @(30, 42)[$index - 1]
    }
    $values = @('Activity', 'Date', 'Community meeting', 'Saturday')
    for ($rowIndex = 1; $rowIndex -le 2; $rowIndex++) {
        for ($columnIndex = 1; $columnIndex -le 2; $columnIndex++) {
            $cells = $table.Cells($rowIndex, $columnIndex, $rowIndex, $columnIndex)
            $cell = $cells.Item(1)
            $fill = $cell.Fill
            $objects.Add($cells)
            $objects.Add($cell)
            $objects.Add($fill)
            $cell.MarginLeft = 0
            $cell.MarginTop = 0
            $cell.MarginRight = 0
            $cell.MarginBottom = 0
            $fill.Visible = 0
            $value = $values[($rowIndex - 1) * 2 + $columnIndex - 1]
            $cell.TextRange.Text = $value
            $range = $cell.TextRange
            $objects.Add($range)
            Set-NativePairedCompositionTextStyle -Range $range -Bold ($rowIndex -eq 1)
            $expectedText.Add($value)
        }
    }
    $wrap = $calendar.TextWrap
    $objects.Add($wrap)
    $wrap.Type = 0
    $expectedShapes.Add([pscustomobject]@{
        name = $calendar.Name
        kind = 'Table'
        pageIndex = 1
        left = 36
        top = 420
        width = 330
        height = 72
        rows = 2
        columns = 2
        cells = $values
    })

    Set-NativePhase 'PairedBrochure-AuthorSecondPage'
    $columnText = ('Small projects help neighbors share skills and improve shared spaces. ' * 10).TrimEnd()
    $columnsShape = Add-NativePairedCompositionTextFrame -Page $second -Name 'PairedBrochureColumns' -Text $columnText -Bounds @(36, 420, 540, 240)
    $frame = $columnsShape.TextFrame
    $objects.Add($frame)
    $frame.Columns = 2
    $frame.ColumnSpacing = 9 # Qualified native half-gap convention: model gap 18 points.
    $frame.MarginLeft = 6
    $frame.MarginTop = 9
    $frame.MarginRight = 12
    $frame.MarginBottom = 3
    $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName frameLayout -NotePropertyValue ([pscustomobject]@{
        columns = 2
        columnSpacing = 9
        modelColumnGapPoints = 18
        left = 6
        top = 9
        right = 12
        bottom = 3
        verticalAlignment = 0
    })
    $image = $second.Shapes.AddPicture($PicturePath, 0, -1, 390, 180, 144, 96)
    $image.Name = 'PairedBrochurePhysicalImage'
    $image.Line.Visible = 0
    $wrap = $image.TextWrap
    $objects.Add($image)
    $objects.Add($wrap)
    $wrap.Type = 0
    $expectedShapes.Add([pscustomobject]@{
        name = $image.Name
        kind = 'Image'
        pageIndex = 2
        left = 390
        top = 180
        width = 144
        height = 96
    })

    Set-NativePhase 'PairedBrochure-AuthorThirdPage'
    $null = Add-NativePairedCompositionTextFrame -Page $third -Name 'PairedBrochureClosing' -Text 'Thank you for taking part. Send your ideas for the next issue.' -Bounds @(36, 420, 540, 72)
    $options = $application.Options
    $objects.Add($options)
    $leadStoryObservations = [pscustomobject]@{
        schema = 'officeimo.publisher.native-multi-paragraph-observations.v1'
        pageCount = 3
        linked = $true
        firstFrameName = 'PairedBrochureLead'
        expectedText = $text
        expectedUtf16Length = 1359
        expectedCrIndices = @(45, 46, 1359)
        expectedLineHeightsPoints = $lineHeights
        expectedFramePageOrder = @(1, 3, 2)
        applicationAutomaticHyphenation = [bool]$options.AutoHyphenate
        applicationOptionsChanged = $false
        candidateStoryAutomaticHyphenation = $true
        nativeStoryAutomaticHyphenation = $null
        nativeStoryAutomaticHyphenationStatus = 'Unassessed: decode the saved native story flag separately; no verified COM setter is used.'
        observations = $observations
    }
    return [pscustomobject]@{
        pages = $layout.pages
        panelWidthPoints = $layout.panelWidthPoints
        masterAssignments = $layout.masterAssignments
        leadStoryObservations = $leadStoryObservations
        controlContract = [pscustomobject]@{
            schema = 'officeimo.publisher.native-paired-composition.v1'
            baseCase = 'TwoPageMaster'
            baseMarkersChanged = $false
            expectedPhysicalMasterDefinitionOrder = @(0, 1, 2)
            expectedMasterShapeCounts = @(3, 3, 3)
            expectedMasterLogoLeftPoints = @(504, 72, 504)
            masterOwnershipObservations = [Collections.Generic.List[object]]::new()
            expectedStoryFramePageOrder = @(1, 3, 2)
            newOrdinaryTextLanguageId = 1033
            newOrdinaryParagraphLineHeightPoints = 16
            candidateNewOrdinaryStoryAutomaticHyphenation = $true
            nativeNewOrdinaryStoryAutomaticHyphenation = $null
            nativeTableStoryAutomaticHyphenation = $null
            nativeStoryHyphenationStatus = 'Unassessed: saved ordinary and table story flags require separate decoding; no COM setter is used.'
            applicationAutomaticHyphenation = [bool]$options.AutoHyphenate
            applicationOptionsChanged = $false
            tableAutomaticHyphenationSetterApplied = $false
            embeddedPicturePlacements = 4
            sourceImageFiles = 1
            nativeImageStoreSharingStatus = 'Unassessed: four native AddPicture calls use one source PNG; saved image-store sharing is a separate check.'
        }
    }
}
