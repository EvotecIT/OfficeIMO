function Get-NativeSizedCompositionProfile {
    <#
    .SYNOPSIS
    Returns the opt-in sized-composition inputs without native operations.
    .PARAMETER Case
    Selects one qualification input; dimensions remain scalar EMU data.
    .EXAMPLE
    Get-NativeSizedCompositionProfile -Case SizedCompositionA4
    .OUTPUTS
    Requested dimensions, exact text and the ten semantic object placements.
    #>
    param(
        [ValidateSet('SizedCompositionA4', 'SizedCompositionPortrait', 'SizedCompositionLandscape', 'SizedCompositionUnassigned')]
        [string] $Case
    )

    $dimensions = switch ($Case) {
        'SizedCompositionA4' { @(7560000, 10692000) }
        'SizedCompositionPortrait' { @(3657600, 5486400) }
        'SizedCompositionLandscape' { @(5486400, 3657600) }
        'SizedCompositionUnassigned' { @(3657600, 5486400) }
    }
    $w = [double]$dimensions[0] / 12700
    $h = [double]$dimensions[1] / 12700
    $linkedText = 'Neighbors share practical ideas, local news and plans for our shared spaces. Join the next meeting to help choose activities for the coming season.'
    $columnText = ('Share your ideas. Meet your neighbors. Plan the next community project. ' * 2).TrimEnd()
    if ($linkedText.Length -ne 147) {
        throw 'The sized-composition linked story requires 147 printable characters and one final native CR.'
    }
    $profiles = @(
        @{
            Name = 'MasterCaption'
            Owner = 'Master'
            Kind = 'TextFrame'
            Type = 17
            Bounds = @(18, 12, ($w - 36), 14)
            Text = 'Community update'
            Size = 8
            LineHeight = 12
            Bold = $false
        },
        @{
            Name = 'MasterBlueBar'
            Owner = 'Master'
            Kind = 'Rectangle'
            Type = 1
            Bounds = @(18, 30, ($w - 36), 3)
            Rgb = 16711680
        },
        @{
            Name = 'MasterLogo'
            Owner = 'Master'
            Kind = 'Image'
            Type = 13
            Bounds = @(($w - 54), ($h - 42), 36, 24)
        },
        @{
            Name = 'Page1Heading'
            Owner = 'Physical1'
            Kind = 'TextFrame'
            Type = 17
            Bounds = @(18, 48, ($w - 36), 30)
            Text = 'Events and projects'
            Size = 16
            LineHeight = 20
            Bold = $true
        },
        @{
            Name = 'LinkedLead'
            Owner = 'Physical1'
            Kind = 'TextFrame'
            Type = 17
            Bounds = @(18, 90, ($w - 36), 14)
            Text = $linkedText
            Size = 10
            LineHeight = 14
            Bold = $false
        },
        @{
            Name = 'LinkedContinuation'
            Owner = 'Physical2'
            Kind = 'TextFrame'
            Type = 17
            Bounds = @(18, 48, ($w - 36), 56)
            Text = ''
            Size = 10
            LineHeight = 14
            Bold = $false
        },
        @{
            Name = 'Calendar'
            Owner = 'Physical1'
            Kind = 'Table'
            Type = 18
            Bounds = @(18, ($h - 96), ($w - 36), 36)
        },
        @{
            Name = 'Page2Columns'
            Owner = 'Physical2'
            Kind = 'TextFrame'
            Type = 17
            Bounds = @(18, 120, ($w - 36), 48)
            Text = $columnText
            Size = 9
            LineHeight = 12
            Bold = $false
        },
        @{
            Name = 'Page1Picture'
            Owner = 'Physical1'
            Kind = 'Image'
            Type = 13
            Bounds = @(($w - 72), 132, 54, 36)
        },
        @{
            Name = 'Page2RedRectangle'
            Owner = 'Physical2'
            Kind = 'Rectangle'
            Type = 1
            Bounds = @(($w - 54), ($h - 54), 36, 12)
            Rgb = 255
        }
    )
    return [pscustomobject]@{
        case = $Case
        widthEmu = $dimensions[0]
        heightEmu = $dimensions[1]
        widthPoints = $w
        heightPoints = $h
        masterName = 'Community update'
        linkedText = $linkedText
        nativeLinkedText = $linkedText + "`r"
        nativeLinkedLength = 148
        columnText = $columnText
        columnWidths = @(84, ($w - 120))
        rowHeights = @(18, 18)
        cellText = @('Activity', 'Date', 'Meeting', 'Saturday')
        objects = $profiles
    }
}

function Set-NativeSizedCompositionTextStyle {
    <#
    .SYNOPSIS
    Formats a freshly reacquired whole story or cell, including its final CR.
    .DESCRIPTION
    Uses explicit ordinary Arial paragraphs and English character language.
    Hyphenation remains a separately observed native default; no options or
    story-hyphenation setter is used.
    #>
    param($Range, [double] $Size, [double] $LineHeight, [bool] $Bold = $false)

    $font = $Range.Font
    $color = $font.Color
    $paragraph = $Range.ParagraphFormat
    $objects.Add($font)
    $objects.Add($color)
    $objects.Add($paragraph)
    $font.Name = 'Arial'
    $font.Size = $Size
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
    $paragraph.SetLineSpacing(4, $LineHeight)
}

function Assert-NativeSizedCompositionControl {
    <#
    .SYNOPSIS
    Verifies actual sized-control ownership, geometry, assignments and content.
    .DESCRIPTION
    Enumerates the native master and physical Shapes collections. Checks native
    getter dimensions separately from encoded header integers, which need file
    decoding. Reacquires complete linked stories and checks actual table tracks
    and cells. Uses a 0.001-point tolerance for native numeric getters.
    .PARAMETER Document
    The task-owned sized composition after authoring or save.
    .PARAMETER Profile
    The scalar input and placement contract returned by Get-NativeSizedCompositionProfile.
    .EXAMPLE
    Assert-NativeSizedCompositionControl -Document $document -Profile $sizedCompositionProfile
    .OUTPUTS
    Serializable actual ownership, dimensions, linked-story and table observations.
    #>
    param($Document, $Profile)

    $setup = $Document.PageSetup
    $pages = $Document.Pages
    $masters = $Document.MasterPages
    $objects.Add($setup)
    $objects.Add($pages)
    $objects.Add($masters)
    $nativeWidth = [double]$setup.PageWidth
    $nativeHeight = [double]$setup.PageHeight
    if ([Math]::Abs($nativeWidth - $Profile.widthPoints) -gt 0.001 -or
        [Math]::Abs($nativeHeight - $Profile.heightPoints) -gt 0.001) {
        throw 'Native sized-composition PageSetup getters differ from the requested point dimensions.'
    }
    if ($masters.Count -ne 1 -or $pages.Count -ne 2) {
        throw 'The sized control requires one centered master and two physical pages.'
    }
    $master = $masters.Item(1)
    $objects.Add($master)
    if ($master.IsTwoPageMaster -or [string]$master.Name -cne $Profile.masterName) {
        throw 'The sized control requires its named single-page native master.'
    }
    $owners = [ordered]@{
        Master = $master
        Physical1 = $pages.Item(1)
        Physical2 = $pages.Item(2)
    }
    $shapeIds = [Collections.Generic.HashSet[int]]::new()
    $ownerObservations = [Collections.Generic.List[object]]::new()
    $linkedFrames = @{}
    foreach ($ownerName in $owners.Keys) {
        $owner = $owners[$ownerName]
        $shapes = $owner.Shapes
        $objects.Add($owner)
        $objects.Add($shapes)
        $expected = @($Profile.objects | Where-Object { $_.Owner -ceq $ownerName })
        if ($shapes.Count -ne $expected.Count) {
            throw ('Unexpected actual sized-control object count on ' + $ownerName + ': ' + $shapes.Count)
        }
        $assignmentId = $null
        if ($ownerName -cne 'Master') {
            $assignedMaster = $owner.Master
            $objects.Add($assignedMaster)
            $assignmentId = [int]$assignedMaster.PageID
            if ($Profile.case -ceq 'SizedCompositionUnassigned') {
                $expectedIgnore = $ownerName -ceq 'Physical2'
                if ($assignmentId -ne [int]$master.PageID -or [bool]$owner.IgnoreMaster -ne $expectedIgnore) {
                    throw ('The unassigned control requires its selected master and exact visibility state on ' + $ownerName)
                }
            } elseif ($assignmentId -ne [int]$master.PageID -or $owner.IgnoreMaster) {
                throw ('The native sized-control master is not applied to ' + $ownerName)
            }
        }
        $seenNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        $shapeObservations = [Collections.Generic.List[object]]::new()
        for ($index = 1; $index -le $shapes.Count; $index++) {
            $shape = $shapes.Item($index)
            $objects.Add($shape)
            $name = [string]$shape.Name
            $matches = @($expected | Where-Object { $_.Name -ceq $name })
            if ($matches.Count -ne 1 -or -not $seenNames.Add($name)) {
                throw ('Unexpected or duplicate sized-control shape on ' + $ownerName + ': ' + $name)
            }
            $wanted = $matches[0]
            $id = [int]$shape.ID
            if ($id -le 0 -or -not $shapeIds.Add($id) -or [int]$shape.Type -ne $wanted.Type) {
                throw ('Unexpected sized-control native shape identity or type: ' + $name)
            }
            $properties = @('Left', 'Top', 'Width', 'Height')
            for ($coordinate = 0; $coordinate -lt $properties.Count; $coordinate++) {
                $property = $properties[$coordinate]
                if ([Math]::Abs([double]$shape.$property - $wanted.Bounds[$coordinate]) -gt 0.001) {
                    throw ('Unexpected sized-control ' + $property + ': ' + $name)
                }
            }
            $wrap = $shape.TextWrap
            $objects.Add($wrap)
            if ([int]$wrap.Type -ne 0) {
                throw ('The sized-control shape requires wrapping disabled: ' + $name)
            }
            if ($Profile.case -ceq 'SizedCompositionUnassigned' -and $wanted.Type -eq 13) {
                $picture = $shape.PictureFormat
                $objects.Add($picture)
                if ($picture.IsEmpty -ne 0 -or $picture.IsLinked -ne 0) {
                    throw ('The unassigned control requires its existing embedded picture: ' + $name)
                }
                foreach ($crop in @('CropLeft', 'CropTop', 'CropRight', 'CropBottom')) {
                    if ([double]$picture.$crop -ne 0) {
                        throw ('The unassigned control requires uncropped picture geometry: ' + $name)
                    }
                }
            }
            $textObservation = $null
            $tableObservation = $null
            if ($wanted.Type -eq 17) {
                $frame = $shape.TextFrame
                $story = $frame.Story
                $whole = $story.TextRange
                $visible = $frame.TextRange
                $font = $whole.Font
                $color = $font.Color
                $paragraph = $whole.ParagraphFormat
                foreach ($wrapper in @($frame, $story, $whole, $visible, $font, $color, $paragraph)) {
                    $objects.Add($wrapper)
                }
                $linked = $name -cin @('LinkedLead', 'LinkedContinuation')
                $expectedText = if ($linked) { $Profile.nativeLinkedText } else { $wanted.Text + "`r" }
                $expectedBold = if ($wanted.Bold) { -1 } else { 0 }
                if ([string]$whole.Text -cne $expectedText -or [int]$whole.Length -ne $expectedText.Length -or
                    $whole.ParagraphsCount -ne 1 -or [string]$font.Name -cne 'Arial' -or
                    [Math]::Abs([double]$font.Size - $wanted.Size) -gt 0.001 -or
                    [int]$font.Bold -ne $expectedBold -or [int]$font.Italic -ne 0 -or
                    [int]$color.RGB -ne 0 -or [int]$whole.LanguageID -ne 1033 -or
                    [int]$paragraph.LineSpacingRule -ne 4 -or
                    [Math]::Abs([double]$paragraph.LineSpacing - $wanted.LineHeight) -gt 0.001) {
                    throw ('The sized-control complete story or explicit text style changed: ' + $name)
                }
                $columnFrame = $name -ceq 'Page2Columns'
                $expectedColumns = if ($columnFrame) { 2 } else { 1 }
                $expectedGap = if ($columnFrame) { 6 } else { 0 }
                $expectedInset = if ($columnFrame) { 3 } else { 0 }
                if ([int]$frame.Columns -ne $expectedColumns -or [int]$frame.VerticalTextAlignment -ne 0 -or
                    [Math]::Abs([double]$frame.ColumnSpacing - $expectedGap) -gt 0.001) {
                    throw ('The sized-control frame column layout changed: ' + $name)
                }
                foreach ($property in @('MarginLeft', 'MarginTop', 'MarginRight', 'MarginBottom')) {
                    if ([Math]::Abs([double]$frame.$property - $expectedInset) -gt 0.001) {
                        throw ('The sized-control frame inset changed: ' + $name + '/' + $property)
                    }
                }
                $hasNext = $frame.HasNextLink -eq -1
                $hasPrevious = $frame.HasPreviousLink -eq -1
                if ($hasNext -ne ($name -ceq 'LinkedLead') -or $hasPrevious -ne ($name -ceq 'LinkedContinuation')) {
                    throw ('The sized-control text-frame link flags changed: ' + $name)
                }
                if ($linked) {
                    $linkedFrames[$name] = @{
                        Frame = $frame
                        ShapeId = $id
                    }
                }
                $textObservation = [pscustomobject]@{
                    wholeStoryText = [string]$whole.Text
                    wholeStoryLength = [int]$whole.Length
                    visibleFrameText = [string]$visible.Text
                    font = [string]$font.Name
                    size = [double]$font.Size
                    languageId = [int]$whole.LanguageID
                    lineHeightPoints = [double]$paragraph.LineSpacing
                    columns = [int]$frame.Columns
                    nativeColumnSpacing = [double]$frame.ColumnSpacing
                    hasNextLink = $hasNext
                    hasPreviousLink = $hasPrevious
                }
            } elseif ($wanted.Type -eq 18) {
                $geometry = Get-NativeTableGeometry $shape
                if ($geometry.columnWidthsPoints.Count -ne 2 -or $geometry.rowHeightsPoints.Count -ne 2) {
                    throw 'The native sized-control calendar requires two rows and two columns.'
                }
                foreach ($track in 0..1) {
                    if ([Math]::Abs($geometry.columnWidthsPoints[$track] - $Profile.columnWidths[$track]) -gt 0.001 -or
                        [Math]::Abs($geometry.rowHeightsPoints[$track] - $Profile.rowHeights[$track]) -gt 0.001) {
                        throw 'The native sized-control calendar track dimensions changed.'
                    }
                }
                $cells = @(Get-NativeTableCellInventory -Shape $shape -Objects $objects)
                if ($cells.Count -ne 4) {
                    throw 'The native sized-control calendar requires four unmerged cells.'
                }
                $nativeTable = $shape.Table
                $objects.Add($nativeTable)
                for ($cellIndex = 0; $cellIndex -lt 4; $cellIndex++) {
                    if ($cells[$cellIndex].text -cne ($Profile.cellText[$cellIndex] + "`r") -or
                        $cells[$cellIndex].columnSpan -ne 1 -or $cells[$cellIndex].rowSpan -ne 1) {
                        throw ('The native sized-control calendar cell changed: ' + $cellIndex)
                    }
                    $rowIndex = if ($cellIndex -lt 2) { 1 } else { 2 }
                    $columnIndex = ($cellIndex % 2) + 1
                    $cellRange = $nativeTable.Cells($rowIndex, $columnIndex, $rowIndex, $columnIndex)
                    $cell = $cellRange.Item(1)
                    $range = $cell.TextRange
                    $font = $range.Font
                    $color = $font.Color
                    $paragraph = $range.ParagraphFormat
                    foreach ($wrapper in @($cellRange, $cell, $range, $font, $color, $paragraph)) {
                        $objects.Add($wrapper)
                    }
                    $expectedBold = if ($rowIndex -eq 1) { -1 } else { 0 }
                    if ([string]$font.Name -cne 'Arial' -or [Math]::Abs([double]$font.Size - 9) -gt 0.001 -or
                        [int]$font.Bold -ne $expectedBold -or [int]$font.Italic -ne 0 -or [int]$color.RGB -ne 0 -or
                        [int]$range.LanguageID -ne 1033 -or [int]$paragraph.LineSpacingRule -ne 4 -or
                        [Math]::Abs([double]$paragraph.LineSpacing - 12) -gt 0.001) {
                        throw ('The native sized-control calendar typography changed: ' + $cellIndex)
                    }
                    $cells[$cellIndex] | Add-Member -NotePropertyName typography -NotePropertyValue ([pscustomobject]@{
                        font = [string]$font.Name
                        size = [double]$font.Size
                        bold = [int]$font.Bold
                        italic = [int]$font.Italic
                        rgb = [int]$color.RGB
                        languageId = [int]$range.LanguageID
                        lineSpacingRule = [int]$paragraph.LineSpacingRule
                        lineHeightPoints = [double]$paragraph.LineSpacing
                    })
                }
                $tableObservation = [pscustomobject]@{
                    geometry = $geometry
                    cells = $cells
                }
            }
            $shapeObservations.Add([pscustomobject]@{
                nativeShapeId = $id
                name = $name
                type = [int]$shape.Type
                left = [double]$shape.Left
                top = [double]$shape.Top
                width = [double]$shape.Width
                height = [double]$shape.Height
                text = $textObservation
                table = $tableObservation
            })
        }
        $ownerObservations.Add([pscustomobject]@{
            role = $ownerName
            pageId = [int]$owner.PageID
            masterPageId = $assignmentId
            shapeCount = [int]$shapes.Count
            shapes = $shapeObservations.ToArray()
        })
        if ($Profile.case -ceq 'SizedCompositionUnassigned' -and $ownerName -cne 'Master') {
            $observation = $ownerObservations[$ownerObservations.Count - 1]
            $observation | Add-Member -NotePropertyName selectedMasterPageId -NotePropertyValue $assignmentId
            $observation | Add-Member -NotePropertyName ignoreMaster -NotePropertyValue ([bool]$owner.IgnoreMaster)
            $observation | Add-Member -NotePropertyName effectiveMasterPageId -NotePropertyValue $(if ($owner.IgnoreMaster) { $null } else { $assignmentId })
        }
    }
    if ($linkedFrames.Count -ne 2) {
        throw 'The sized-control story requires exactly two linked frames.'
    }
    $next = $linkedFrames.LinkedLead.Frame.NextLinkedTextFrame
    $previous = $linkedFrames.LinkedContinuation.Frame.PreviousLinkedTextFrame
    $nextParent = $next.Parent
    $previousParent = $previous.Parent
    foreach ($wrapper in @($next, $previous, $nextParent, $previousParent)) {
        $objects.Add($wrapper)
    }
    if ([int]$nextParent.ID -ne $linkedFrames.LinkedContinuation.ShapeId -or
        [int]$previousParent.ID -ne $linkedFrames.LinkedLead.ShapeId) {
        throw 'The sized-control story does not retain reciprocal page 1/2 links.'
    }
    return [pscustomobject]@{
        schema = 'officeimo.publisher.native-sized-composition-observation.v1'
        status = 'Observed'
        requestedWidthEmu = $Profile.widthEmu
        requestedHeightEmu = $Profile.heightEmu
        nativeGetterWidthPoints = $nativeWidth
        nativeGetterHeightPoints = $nativeHeight
        getterWidthDeltaPoints = $nativeWidth - $Profile.widthPoints
        getterHeightDeltaPoints = $nativeHeight - $Profile.heightPoints
        nativeHeaderDimensionsStatus = 'Unassessed: native getter agreement does not establish exact encoded header integers.'
        coordinateTolerancePoints = 0.001
        masterIsTwoPage = [bool]$master.IsTwoPageMaster
        uniqueNativeShapeCount = $shapeIds.Count
        linkedStoryLengthIncludingCR = $Profile.nativeLinkedLength
        linkedFramePageOrder = @(1, 2)
        owners = $ownerObservations.ToArray()
    }
}

function Add-NativeSizedCompositionControl {
    <#
    .SYNOPSIS
    Authors the independent mixed-content control at requested EMU dimensions.
    .DESCRIPTION
    Sets page dimensions before adding artwork, then uses one centered master
    and two physical pages. Reuses the worker's text and synthetic PNG helpers.
    One PNG is embedded twice; saved image-store sharing is assessed separately.
    Application hyphenation is observed only, and table hyphenation is not set.
    .PARAMETER Document
    The isolated worker's freshly created task-owned publication.
    .PARAMETER FirstPage
    Its first physical page, before this helper adds page two.
    .PARAMETER Profile
    The requested scalar dimensions and placements from Get-NativeSizedCompositionProfile.
    .PARAMETER PicturePath
    Unused task-owned path for the existing helper's synthetic 32-pixel PNG.
    .EXAMPLE
    Add-NativeSizedCompositionControl -Document $document -FirstPage $page -Profile $sizedCompositionProfile -PicturePath $picturePath
    .OUTPUTS
    Page/master assignments and the sized-control input and observation contract.
    #>
    param($Document, $FirstPage, $Profile, [string] $PicturePath)

    Set-NativePhase 'SizedComposition-SetPageDimensions'
    $setup = $Document.PageSetup
    $pages = $Document.Pages
    $masters = $Document.MasterPages
    foreach ($wrapper in @($setup, $pages, $masters)) {
        $objects.Add($wrapper)
    }
    $setup.PageWidth = $Profile.widthPoints
    $setup.PageHeight = $Profile.heightPoints
    if ($pages.Count -ne 1 -or $masters.Count -ne 1) {
        throw 'The sized composition requires the untouched one-page, one-master native publication.'
    }
    $master = $masters.Item(1)
    $objects.Add($master)
    if ($master.IsTwoPageMaster) {
        throw 'The sized composition requires a centered native master.'
    }
    $master.Name = $Profile.masterName
    $second = $pages.Add(1, 1)
    $objects.Add($second)
    $FirstPage.Master = $master
    $second.Master = $master
    if ($Profile.case -ceq 'SizedCompositionUnassigned') {
        Set-NativePhase 'SizedComposition-SuppressSecondPageMaster'
        $FirstPage.IgnoreMaster = $false
        $second.IgnoreMaster = $true
    }
    $owners = @{
        Master = $master
        Physical1 = $FirstPage
        Physical2 = $second
    }
    $frames = @{}
    Set-NativePhase 'SizedComposition-CreateReferencePicture'
    New-ReferencePicture $PicturePath
    foreach ($profileObject in $Profile.objects) {
        Set-NativePhase ('SizedComposition-Author-' + $profileObject.Name)
        $owner = $owners[$profileObject.Owner]
        $bounds = $profileObject.Bounds
        if ($profileObject.Type -eq 17) {
            $shape = Add-TextFrame $owner $profileObject.Name $profileObject.Text $bounds[0] $bounds[1] $bounds[2] $bounds[3] $profileObject.Size
            $frame = $shape.TextFrame
            $story = $frame.Story
            $whole = $story.TextRange
            foreach ($wrapper in @($frame, $story, $whole)) {
                $objects.Add($wrapper)
            }
            Set-NativeSizedCompositionTextStyle -Range $whole -Size $profileObject.Size -LineHeight $profileObject.LineHeight -Bold $profileObject.Bold
            $columnFrame = $profileObject.Name -ceq 'Page2Columns'
            $frame.Columns = if ($columnFrame) { 2 } else { 1 }
            $frame.ColumnSpacing = if ($columnFrame) { 6 } else { 0 } # Native half-gap for full model gap 12.
            $frame.VerticalTextAlignment = 0
            if ($columnFrame) {
                $frame.MarginLeft = 3
                $frame.MarginTop = 3
                $frame.MarginRight = 3
                $frame.MarginBottom = 3
            }
            $frames[$profileObject.Name] = $frame
        } elseif ($profileObject.Type -eq 1) {
            $shape = $owner.Shapes.AddShape(1, $bounds[0], $bounds[1], $bounds[2], $bounds[3])
            $shape.Fill.ForeColor.RGB = $profileObject.Rgb
        } elseif ($profileObject.Type -eq 13) {
            $shape = $owner.Shapes.AddPicture($PicturePath, 0, -1, $bounds[0], $bounds[1], $bounds[2], $bounds[3])
        } else {
            $shape = $owner.Shapes.AddTable(2, 2, $bounds[0], $bounds[1], $bounds[2], $bounds[3])
            $table = $shape.Table
            $columns = $table.Columns
            $rows = $table.Rows
            foreach ($wrapper in @($table, $columns, $rows)) {
                $objects.Add($wrapper)
            }
            $table.ApplyAutoFormat(-2, $false, $false, $true, $true)
            foreach ($track in 1..2) {
                $column = $columns.Item($track)
                $row = $rows.Item($track)
                $objects.Add($column)
                $objects.Add($row)
                # Column.Width is VARIANT. Keep both values Double at this COM
                # assignment site, preserving fractional points without mixing
                # the first Int32 width with a later boxed Double.
                $columnWidth = [double]$Profile.columnWidths[$track - 1]
                $column.Width = $columnWidth
                $row.Height = $Profile.rowHeights[$track - 1]
            }
            for ($rowIndex = 1; $rowIndex -le 2; $rowIndex++) {
                for ($columnIndex = 1; $columnIndex -le 2; $columnIndex++) {
                    $cells = $table.Cells($rowIndex, $columnIndex, $rowIndex, $columnIndex)
                    $cell = $cells.Item(1)
                    $fill = $cell.Fill
                    foreach ($wrapper in @($cells, $cell, $fill)) {
                        $objects.Add($wrapper)
                    }
                    $cell.MarginLeft = 0
                    $cell.MarginTop = 0
                    $cell.MarginRight = 0
                    $cell.MarginBottom = 0
                    $fill.Visible = 0
                    $value = $Profile.cellText[($rowIndex - 1) * 2 + $columnIndex - 1]
                    $cell.TextRange.Text = $value
                    $range = $cell.TextRange
                    $objects.Add($range)
                    Set-NativeSizedCompositionTextStyle -Range $range -Size 9 -LineHeight 12 -Bold ($rowIndex -eq 1)
                    $expectedText.Add($value)
                }
            }
        }
        $shape.Name = $profileObject.Name
        if ($profileObject.Type -ne 18) {
            $shape.Line.Visible = 0
        }
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        $wrap.Type = 0
        if ($profileObject.Type -ne 17) {
            $objects.Add($shape)
            $expectedShapes.Add([pscustomobject]@{
                name = $profileObject.Name
                kind = $profileObject.Kind
                pageRole = if ($profileObject.Owner -ceq 'Master') { 'Master' } else { 'Physical' }
                pageIndex = if ($profileObject.Owner -ceq 'Physical2') { 2 } else { 1 }
                left = $bounds[0]
                top = $bounds[1]
                width = $bounds[2]
                height = $bounds[3]
            })
        }
    }
    Set-NativePhase 'SizedComposition-LinkStory'
    $story = $frames.LinkedLead.Story
    $whole = $story.TextRange
    $objects.Add($story)
    $objects.Add($whole)
    if ([string]$whole.Text -cne $Profile.nativeLinkedText -or [int]$whole.Length -ne 148 -or $whole.ParagraphsCount -ne 1) {
        throw 'The complete unlinked sized-control story must retain its 147 characters and final CR.'
    }
    $frames.LinkedLead.NextLinkedTextFrame = $frames.LinkedContinuation
    $options = $application.Options
    $objects.Add($options)
    $layout = [pscustomobject]@{
        pages = 2
        panelWidthPoints = $null
        masterAssignments = @(
            [pscustomobject]@{
                pageIndex = 1
                masterPageId = [int]$master.PageID
            },
            [pscustomobject]@{
                pageIndex = 2
                masterPageId = [int]$master.PageID
            }
        )
        controlContract = [pscustomobject]@{
            schema = 'officeimo.publisher.native-sized-composition.v1'
            case = $Profile.case
            requestedWidthEmu = $Profile.widthEmu
            requestedHeightEmu = $Profile.heightEmu
            requestedWidthPoints = $Profile.widthPoints
            requestedHeightPoints = $Profile.heightPoints
            expectedOwnerShapeCounts = @(3, 4, 3)
            expectedSemanticObjectCount = 10
            expectedTextFrameCount = 5
            expectedRectangleCount = 2
            expectedPictureCount = 2
            expectedTableCount = 1
            expectedStoriesIncludingTable = 5
            expectedSharedModelImageAssets = 1
            expectedLinkedStoryLengthIncludingCR = 148
            expectedLinkedFramePageOrder = @(1, 2)
            applicationAutomaticHyphenation = [bool]$options.AutoHyphenate
            applicationOptionsChanged = $false
            candidateOrdinaryStoryAutomaticHyphenation = $true
            nativeOrdinaryStoryAutomaticHyphenation = $null
            nativeTableStoryAutomaticHyphenation = $null
            nativeStoryHyphenationStatus = 'Unassessed: decode saved story flags separately; no COM setter is used.'
            tableAutomaticHyphenationSetterApplied = $false
            sourceImageFiles = 1
            embeddedPicturePlacements = 2
            nativeImageStoreSharingStatus = 'Unassessed: two AddPicture calls use one source PNG; saved BStore sharing and reader asset byte identity need separate checks.'
            observations = [Collections.Generic.List[object]]::new()
        }
    }
    if ($Profile.case -ceq 'SizedCompositionUnassigned') {
        foreach ($assignment in $layout.masterAssignments) {
            $ignored = $assignment.pageIndex -eq 2
            $assignment | Add-Member -NotePropertyName selectedMasterPageId -NotePropertyValue $assignment.masterPageId
            $assignment | Add-Member -NotePropertyName ignoreMaster -NotePropertyValue $ignored
            $assignment | Add-Member -NotePropertyName effectiveMasterPageId -NotePropertyValue $(if ($ignored) { $null } else { $assignment.masterPageId })
        }
        $layout.controlContract | Add-Member -NotePropertyName masterAssignmentSemantics -NotePropertyValue 'masterPageId and selectedMasterPageId name the retained native selection; IgnoreMaster suppresses it on page 2, whose effectiveMasterPageId is null.'
        $layout.controlContract | Add-Member -NotePropertyName expectedMasterAssignments -NotePropertyValue $layout.masterAssignments
    }
    return $layout
}
