function Add-NativeLinkedStoryControl {
    <#
    .SYNOPSIS
    Authors one story flowing through pages one, three and two, plus an independent label.
    .PARAMETER Document
    Task-owned new publication.
    .PARAMETER FirstPage
    Its first physical page.
    .PARAMETER Overflow
    Uses a longer story that exceeds the final frame.
    .OUTPUTS
    The physical page count.
    #>
    param($Document, $FirstPage, [switch]$Overflow)

    $second = $Document.Pages.Add(1, 1)
    $third = $Document.Pages.Add(1, 2)
    $objects.Add($second)
    $objects.Add($third)
    $count = if ($Overflow) { 24 } else { 6 }
    $text = ('Linked stories preserve reading order across publication pages. ' * $count).TrimEnd()
    $first = Add-TextFrame $FirstPage 'StoryStart' $text 72 72 216 72 12
    $label = Add-TextFrame $FirstPage 'IndependentLabel' 'Independent story' 360 72 180 54 14
    $middle = Add-TextFrame $third 'StoryMiddle' '' 72 144 216 72 12
    $last = Add-TextFrame $second 'StoryEnd' '' 72 216 360 108 12
    $first.TextFrame.NextLinkedTextFrame = $middle.TextFrame
    $middle.TextFrame.NextLinkedTextFrame = $last.TextFrame
    foreach ($shape in @($first, $label, $middle, $last)) {
        $shape.TextFrame.ColumnSpacing = 0
        $paragraph = $shape.TextFrame.TextRange.ParagraphFormat
        $objects.Add($paragraph)
        $paragraph.SpaceBefore = 0
        $paragraph.SpaceAfter = 0
        $paragraph.SetLineSpacing(4, 18)
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        $wrap.Type = 0
        if ($wrap.Type -ne 0) { throw 'Linked story control requires wrapping disabled.' }
    }
    if ($first.TextFrame.HasNextLink -ne -1 -or $first.TextFrame.HasPreviousLink -ne 0 -or
        $middle.TextFrame.HasNextLink -ne -1 -or $middle.TextFrame.HasPreviousLink -ne -1 -or
        $last.TextFrame.HasNextLink -ne 0 -or $last.TextFrame.HasPreviousLink -ne -1) {
        throw 'Linked story control did not retain the three-frame chain.'
    }
    return 3
}
