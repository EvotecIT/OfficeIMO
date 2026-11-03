function Get-NativeCaptionRangeObservation {
    <#
    .SYNOPSIS
    Observes one caption's full range and selected characters without editing it.
    .PARAMETER Shape
    The task-owned sole caption text box.
    .PARAMETER Stage
    Names the completed authoring or save operation being observed.
    .OUTPUTS
    Range metadata and first/last visible-character and final-CR observations.
    A missing or inaccessible final CR remains explicitly unassessed.
    #>
    param($Shape, [string] $Stage)

    $frame = $Shape.TextFrame
    $objects.Add($frame)
    $range = $frame.TextRange
    $objects.Add($range)
    $font = $range.Font
    $objects.Add($font)
    $text = [string]$range.Text
    $hasFinalCr = $text.EndsWith("`r", [StringComparison]::Ordinal)
    $visibleLength = $text.Length
    if ($hasFinalCr) {
        $visibleLength--
    }
    $selections = [Collections.Generic.List[object]]::new()
    if ($visibleLength -gt 0) {
        $selections.Add([pscustomobject]@{ role = 'FirstVisible'; index = 1 })
        $selections.Add([pscustomobject]@{ role = 'LastVisible'; index = $visibleLength })
    }
    if ($hasFinalCr) {
        $selections.Add([pscustomobject]@{ role = 'FinalCR'; index = $text.Length })
    }
    $points = [Collections.Generic.List[object]]::new()
    foreach ($selection in $selections) {
        # The installed TextRange interface exposes Characters(Start, Length).
        # Record the returned character as well as its requested one-based index;
        # point identity is evidence, rather than an assumed COM range boundary.
        try {
            $character = $range.Characters($selection.index, 1)
            $objects.Add($character)
            $characterFont = $character.Font
            $objects.Add($characterFont)
            $observedText = [string]$character.Text
            $expectedCharacter = $text.Substring($selection.index - 1, 1)
            $points.Add([pscustomobject]@{
                role = $selection.role
                requestedIndex = $selection.index
                status = if ($observedText -ceq $expectedCharacter) { 'Observed' } else { 'DifferentCharacter' }
                text = $observedText
                expectedCharacter = $expectedCharacter
                start = [int]$character.Start
                end = [int]$character.End
                length = [int]$character.Length
                font = [string]$characterFont.Name
                fontSizePoints = [double]$characterFont.Size
                languageId = [int]$character.LanguageID
            })
        } catch {
            $points.Add([pscustomobject]@{
                role = $selection.role
                requestedIndex = $selection.index
                status = 'Unavailable'
                message = $_.Exception.Message
                hresult = $_.Exception.HResult
            })
        }
    }
    $finalCr = @($points | Where-Object { $_.role -eq 'FinalCR' })
    $finalCrStatus = if (-not $hasFinalCr) {
        'NotMaterializedInRangeText'
    } else {
        $finalCr[0].status
    }
    return [pscustomobject]@{
        stage = $Stage
        text = $text
        utf16Length = $text.Length
        start = [int]$range.Start
        end = [int]$range.End
        length = [int]$range.Length
        font = [string]$font.Name
        fontSizePoints = [double]$font.Size
        languageId = [int]$range.LanguageID
        finalCrStatus = $finalCrStatus
        points = $points.ToArray()
    }
}

function Add-NativeCaptionTypingControl {
    <#
    .SYNOPSIS
    Authors a sole caption with original-range or explicit full-range formatting.
    .DESCRIPTION
    Uses Add-TextFrame unchanged, then optionally formats a reacquired full range.
    Clear/refill observes retained typing defaults without reapplying formatting.
    It does not alter native style libraries, publication defaults or Office settings.
    .PARAMETER FirstPage
    The task-owned new publication page.
    .PARAMETER Mode
    PreRange retains Add-TextFrame's existing formatting. FullRange formats the
    reacquired full range; Explicit also assigns language identifier 1033.
    .PARAMETER Refilled
    Clears the full range and refills a newly acquired range with fixed text.
    .EXAMPLE
    Add-NativeCaptionTypingControl -FirstPage $page -Mode Explicit -Refilled
    .OUTPUTS
    Point and aggregate observations. The caller records the single physical page.
    #>
    param(
        $FirstPage,
        [ValidateSet('PreRange', 'FullRange', 'Explicit')]
        [string] $Mode,
        [switch] $Refilled
    )

    Set-NativePhase 'AuthorCaptionTypingControl'
    $caption = 'Independent table label'
    $shape = Add-TextFrame $FirstPage 'CaptionTyping' $caption 72 72 300 54 12
    $frame = $shape.TextFrame
    $objects.Add($frame)
    $range = $frame.TextRange
    $objects.Add($range)
    $paragraph = $range.ParagraphFormat
    $objects.Add($paragraph)
    $paragraph.SpaceBefore = 0
    $paragraph.SpaceAfter = 0
    $paragraph.SetLineSpacing(4, 18)
    $observations = [Collections.Generic.List[object]]::new()
    $observations.Add((Get-NativeCaptionRangeObservation -Shape $shape -Stage 'AfterTextFrameHelper'))
    if ($Mode -ne 'PreRange') {
        $fullRange = $frame.TextRange
        $objects.Add($fullRange)
        $font = $fullRange.Font
        $objects.Add($font)
        $color = $font.Color
        $objects.Add($color)
        $font.Name = 'Arial'
        $font.Size = 12
        $color.RGB = 0
        if ($Mode -eq 'Explicit') {
            $fullRange.LanguageID = 1033
        }
        $observations.Add((Get-NativeCaptionRangeObservation -Shape $shape -Stage 'AfterFullRangeFormatting'))
    }
    if ($Refilled) {
        Set-NativePhase 'ClearAndRefillCaption'
        $clearRange = $frame.TextRange
        $objects.Add($clearRange)
        $clearRange.Text = ''
        $observations.Add((Get-NativeCaptionRangeObservation -Shape $shape -Stage 'AfterClear'))
        $refillRange = $frame.TextRange
        $objects.Add($refillRange)
        $refillRange.Text = 'Refilled caption'
        $observations.Add((Get-NativeCaptionRangeObservation -Shape $shape -Stage 'AfterRefill'))
        if ($expectedText.Count -ne 1 -or $expectedText[0] -cne $caption) {
            throw 'Caption typing controls require the sole text frame created by this helper.'
        }
        $expectedText[0] = 'Refilled caption'
        $expectedShape = $expectedShapes[$expectedShapes.Count - 1]
        $expectedShape | Add-Member -NotePropertyName initialFormatting -NotePropertyValue ([pscustomobject]@{
            font = $expectedShape.font
            size = $expectedShape.size
        })
        $expectedShape.font = $null
        $expectedShape.size = $null
        $expectedShape | Add-Member -NotePropertyName finalFormatting -NotePropertyValue 'Unqualified: clear/refill does not reapply font or size; captionRangeObservations records the native result.'
    }
    return [pscustomobject]@{
        schema = 'officeimo.publisher.native-caption-typing-observations.v1'
        mode = $Mode
        refilled = [bool]$Refilled
        originalRangeMembership = 'Unassessed: Add-TextFrame retains its original COM range internally and does not expose its identity.'
        pointSelection = 'TextRange.Characters(Start, Length); one-based requested indices are checked against the returned character text.'
        observations = $observations
    }
}
