function Get-NativeMultiParagraphObservation {
    <#
    .SYNOPSIS
    Observes complete story ranges, CR typing points and paragraph layout.
    .DESCRIPTION
    Reacquires TextFrame.Story.TextRange for each observation. Requested point
    indices are checked against returned character text; unavailable observations
    remain explicit instead of implying range membership from a prior COM wrapper.
    All COM wrappers are retained in the isolated caller's objects list.
    .PARAMETER Shape
    The first task-owned text box of the ordinary story.
    .PARAMETER ExpectedText
    The independently authored text, including all three final paragraph CRs.
    .PARAMETER Stage
    Names the completed authoring or save operation being observed.
    .OUTPUTS
    Native frame/story range boundaries, selected character properties and paragraphs.
    #>
    param($Shape, [string] $ExpectedText, [string] $Stage)

    $frame = $Shape.TextFrame
    $objects.Add($frame)
    $frameRange = $frame.TextRange
    $objects.Add($frameRange)
    $story = $frame.Story
    $objects.Add($story)
    $whole = $story.TextRange
    $objects.Add($whole)
    $text = [string]$whole.Text
    $crIndices = [Collections.Generic.List[int]]::new()
    for ($index = 0; $index -lt $text.Length; $index++) {
        if ($text[$index] -eq "`r") {
            $crIndices.Add($index + 1)
        }
    }
    $selections = @(
        @{ Role = 'HeadingFirst'; Index = 1 },
        @{ Role = 'HeadingLast'; Index = 25 },
        @{ Role = 'FirstBodyFirst'; Index = 26 },
        @{ Role = 'FirstBodyLast'; Index = 44 },
        @{ Role = 'FirstParagraphCR'; Index = 45 },
        @{ Role = 'EmptyParagraphCR'; Index = 46 },
        @{ Role = 'FinalPrefixFirst'; Index = 47 },
        @{ Role = 'FinalPrefixLast'; Index = 63 },
        @{ Role = 'FinalBodyFirst'; Index = 64 },
        @{ Role = 'FinalBodyLast'; Index = 1358 },
        @{ Role = 'FinalParagraphCR'; Index = 1359 }
    )
    $points = [Collections.Generic.List[object]]::new()
    foreach ($selection in $selections) {
        try {
            if ($selection.Index -gt $text.Length) {
                throw 'Requested point is outside the currently observed story range.'
            }
            $character = $whole.Characters($selection.Index, 1)
            $objects.Add($character)
            $font = $character.Font
            $objects.Add($font)
            $color = $font.Color
            $objects.Add($color)
            $observedText = [string]$character.Text
            $expectedCharacter = $ExpectedText.Substring($selection.Index - 1, 1)
            $points.Add([pscustomobject]@{
                role = $selection.Role
                requestedIndex = $selection.Index
                status = if ($observedText -ceq $expectedCharacter) { 'Observed' } else { 'DifferentCharacter' }
                text = $observedText
                expectedCharacter = $expectedCharacter
                start = [int]$character.Start
                end = [int]$character.End
                length = [int]$character.Length
                font = [string]$font.Name
                fontSizePoints = [double]$font.Size
                bold = [int]$font.Bold
                rgb = [int]$color.RGB
                languageId = [int]$character.LanguageID
            })
        } catch {
            $points.Add([pscustomobject]@{
                role = $selection.Role
                requestedIndex = $selection.Index
                status = 'Unavailable'
                message = $_.Exception.Message
                hresult = $_.Exception.HResult
            })
        }
    }
    $paragraphs = [Collections.Generic.List[object]]::new()
    foreach ($ordinal in 1..3) {
        try {
            $range = $whole.Paragraphs($ordinal, 1)
            $objects.Add($range)
            $format = $range.ParagraphFormat
            $objects.Add($format)
            $font = $range.Font
            $objects.Add($font)
            $paragraphs.Add([pscustomobject]@{
                ordinal = $ordinal
                status = 'Observed'
                text = [string]$range.Text
                start = [int]$range.Start
                end = [int]$range.End
                length = [int]$range.Length
                font = [string]$font.Name
                fontSizePoints = [double]$font.Size
                bold = [int]$font.Bold
                languageId = [int]$range.LanguageID
                alignment = [int]$format.Alignment
                leftIndent = [double]$format.LeftIndent
                rightIndent = [double]$format.RightIndent
                firstLineIndent = [double]$format.FirstLineIndent
                spaceBefore = [double]$format.SpaceBefore
                spaceAfter = [double]$format.SpaceAfter
                lineSpacingRule = [int]$format.LineSpacingRule
                lineSpacingPoints = [double]$format.LineSpacing
                listType = [int]$format.ListType
            })
        } catch {
            $paragraphs.Add([pscustomobject]@{
                ordinal = $ordinal
                status = 'Unavailable'
                message = $_.Exception.Message
                hresult = $_.Exception.HResult
            })
        }
    }
    return [pscustomobject]@{
        stage = $Stage
        shape = [string]$Shape.Name
        frameRange = [pscustomobject]@{
            text = [string]$frameRange.Text
            start = [int]$frameRange.Start
            end = [int]$frameRange.End
            length = [int]$frameRange.Length
        }
        storyRange = [pscustomobject]@{
            accessor = 'TextFrame.Story.TextRange'
            text = $text
            start = [int]$whole.Start
            end = [int]$whole.End
            length = [int]$whole.Length
            utf16Length = $text.Length
            matchesExpectedText = $text -ceq $ExpectedText
            crIndices = $crIndices.ToArray()
        }
        points = $points.ToArray()
        paragraphs = $paragraphs.ToArray()
    }
}

function Add-NativeMultiParagraphControl {
    <#
    .SYNOPSIS
    Authors three differently styled paragraphs in one frame or a page 1/3/2 chain.
    .DESCRIPTION
    Formats a freshly reacquired complete unlinked story, including the interior,
    empty-paragraph and final CRs, before connecting continuation frames. Native
    observations are reacquired after linking and saving. Application options are
    observed only; the saved native story hyphenation flag needs separate decoding.
    .PARAMETER Document
    The isolated worker's new native publication.
    .PARAMETER FirstPage
    Its first physical page.
    .PARAMETER Linked
    Continues the story on physical pages three and two, in that order.
    .EXAMPLE
    Add-NativeMultiParagraphControl -Document $document -FirstPage $page -Linked
    .OUTPUTS
    The page count, authored contract, hyphenation boundary and native observations.
    #>
    param($Document, $FirstPage, [switch] $Linked)

    Set-NativePhase 'AuthorMultiParagraphControl'
    $heading = 'Multi-paragraph heading. '
    $firstBody = 'Body follows below.'
    $finalPrefix = 'Final paragraph: '
    $finalBody = ('Complete ordinary stories retain every paragraph and all overflow text. ' * 18).TrimEnd()
    $text = $heading + $firstBody + "`r`r" + $finalPrefix + $finalBody + "`r"
    if ($text.Length -ne 1359 -or $text[44] -ne "`r" -or $text[45] -ne "`r" -or $text[1358] -ne "`r") {
        throw 'The independent multi-paragraph story must retain the 45/46/1359 CR endpoints.'
    }
    $pages = $Document.Pages
    $objects.Add($pages)
    $profiles = [Collections.Generic.List[object]]::new()
    $profiles.Add(@{ Page = $FirstPage; PageIndex = 1; Name = 'MultiParagraphStart'; Top = 72; Width = 216; Height = 72 })
    if ($Linked) {
        $second = $pages.Add(1, 1)
        $objects.Add($second)
        $third = $pages.Add(1, 2)
        $objects.Add($third)
        $profiles.Add(@{ Page = $third; PageIndex = 3; Name = 'MultiParagraphMiddle'; Top = 144; Width = 216; Height = 72 })
        $profiles.Add(@{ Page = $second; PageIndex = 2; Name = 'MultiParagraphEnd'; Top = 216; Width = 360; Height = 108 })
    }
    $shapes = [Collections.Generic.List[object]]::new()
    $frames = [Collections.Generic.List[object]]::new()
    foreach ($profile in $profiles) {
        $inputText = if ($shapes.Count -eq 0) { $text.Substring(0, $text.Length - 1) } else { '' }
        $shape = Add-TextFrame $profile.Page $profile.Name $inputText 72 $profile.Top $profile.Width $profile.Height 12
        $frame = $shape.TextFrame
        $objects.Add($frame)
        $frame.Columns = 1
        $frame.ColumnSpacing = 0
        $frame.VerticalTextAlignment = 0
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        $wrap.Type = 0
        $expected = $expectedShapes[$expectedShapes.Count - 1]
        $expected.size = $null
        $expected | Add-Member -NotePropertyName pageIndex -NotePropertyValue $profile.PageIndex
        $shapes.Add($shape)
        $frames.Add($frame)
    }
    $story = $frames[0].Story
    $objects.Add($story)
    $whole = $story.TextRange
    $objects.Add($whole)
    if ([string]$whole.Text -cne $text -or [int]$whole.Length -ne $text.Length) {
        throw 'The reacquired native story range does not contain the complete text and final CR.'
    }
    $font = $whole.Font
    $objects.Add($font)
    $color = $font.Color
    $objects.Add($color)
    $font.Name = 'Arial'
    $font.Size = 12
    $font.Bold = 0
    $font.Italic = 0
    $color.RGB = 0
    $whole.LanguageID = 1033
    foreach ($run in @(
        @{ Start = 1; Length = 25; Size = 16 },
        @{ Start = 46; Length = 1; Size = 14 },
        @{ Start = 47; Length = 17; Size = 12 }
    )) {
        $range = $whole.Characters($run.Start, $run.Length)
        $objects.Add($range)
        $runFont = $range.Font
        $objects.Add($runFont)
        $runFont.Size = $run.Size
        $runFont.Bold = -1
    }
    $lineHeights = @(22, 20, 18)
    foreach ($ordinal in 1..3) {
        $range = $whole.Paragraphs($ordinal, 1)
        $objects.Add($range)
        $format = $range.ParagraphFormat
        $objects.Add($format)
        $format.Alignment = 0
        $format.LeftIndent = 0
        $format.RightIndent = 0
        $format.FirstLineIndent = 0
        $format.SpaceBefore = 0
        $format.SpaceAfter = 0
        $format.SetLineSpacing(4, $lineHeights[$ordinal - 1])
        if ([int]$format.ListType -ne 255) { # PbListType.pbListTypeNone in the installed Publisher type library.
            throw 'The multi-paragraph control requires ordinary paragraphs without list labels.'
        }
    }
    $observations = [Collections.Generic.List[object]]::new()
    $observations.Add((Get-NativeMultiParagraphObservation -Shape $shapes[0] -ExpectedText $text -Stage 'BeforeLinking'))
    if ($Linked) {
        $frames[0].NextLinkedTextFrame = $frames[1]
        $frames[1].NextLinkedTextFrame = $frames[2]
        if ($frames[0].HasNextLink -ne -1 -or $frames[0].HasPreviousLink -ne 0 -or
            $frames[1].HasNextLink -ne -1 -or $frames[1].HasPreviousLink -ne -1 -or
            $frames[2].HasNextLink -ne 0 -or $frames[2].HasPreviousLink -ne -1) {
            throw 'The native multi-paragraph control did not retain the page 1/3/2 frame chain.'
        }
        $observations.Add((Get-NativeMultiParagraphObservation -Shape $shapes[0] -ExpectedText $text -Stage 'AfterLinking'))
    }
    $options = $application.Options
    $objects.Add($options)
    return [pscustomobject]@{
        schema = 'officeimo.publisher.native-multi-paragraph-observations.v1'
        pageCount = if ($Linked) { 3 } else { 1 }
        linked = [bool]$Linked
        firstFrameName = 'MultiParagraphStart'
        expectedText = $text
        expectedUtf16Length = 1359
        expectedCrIndices = @(45, 46, 1359)
        expectedLineHeightsPoints = $lineHeights
        expectedFramePageOrder = if ($Linked) { @(1, 3, 2) } else { @(1) }
        applicationAutomaticHyphenation = [bool]$options.AutoHyphenate
        applicationOptionsChanged = $false
        candidateStoryAutomaticHyphenation = $false
        nativeStoryAutomaticHyphenation = $null
        nativeStoryAutomaticHyphenationStatus = 'Unassessed: decode the saved native story flag separately; no verified COM setter is used.'
        observations = $observations
    }
}
