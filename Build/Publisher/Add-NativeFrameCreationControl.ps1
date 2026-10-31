function Add-NativeFrameCreationControl {
    <#
    .SYNOPSIS
    Authors six column, inset and vertical-alignment frames on two pages.
    .PARAMETER Document
    Task-owned new publication.
    .PARAMETER FirstPage
    Its first physical page.
    .PARAMETER Edited
    Applies a common replacement layout to all six frames.
    .OUTPUTS
    The physical page count.
    #>
    param($Document, $FirstPage, [switch]$Edited)

    $second = $Document.Pages.Add(1, 1)
    $objects.Add($second)
    $profiles = @(
        @{ Columns = 1; Gap = 0; Insets = @(12, 6, 18, 9); Vertical = 0 },
        @{ Columns = 2; Gap = 6; Insets = @(0, 0, 0, 0); Vertical = 0 },
        @{ Columns = 3; Gap = 9; Insets = @(0, 0, 0, 0); Vertical = 0 },
        @{ Columns = 1; Gap = 0; Insets = @(12, 6, 18, 9); Vertical = 1; Short = $true },
        @{ Columns = 1; Gap = 0; Insets = @(0, 0, 0, 0); Vertical = 2; Short = $true },
        @{ Columns = 3; Gap = 6.05; Insets = @(1.21, 2.42, 3.63, 4.84); Vertical = 0 }
    )
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        $owner = if ($index -lt 3) { $FirstPage } else { $second }
        $text = if ($profile.ContainsKey('Short')) { 'Short native body.' } else {
            ('One two three four five six seven eight nine ten. ' * 12).TrimEnd()
        }
        $shape = Add-TextFrame $owner ('Frame' + $index) $text 54 (54 + ($index % 3) * 228) 504 180 12
        $frame = $shape.TextFrame
        $paragraph = $frame.TextRange.ParagraphFormat
        $objects.Add($frame)
        $objects.Add($paragraph)
        $paragraph.SpaceBefore = 0
        $paragraph.SpaceAfter = 0
        $paragraph.SetLineSpacing(4, 18)
        $insets = if ($Edited) { @(9, 3, 15, 6) } else { $profile.Insets }
        $frame.Columns = if ($Edited) { 2 } else { $profile.Columns }
        # Publisher's automation value is half the full model intercolumn gap.
        $frame.ColumnSpacing = [double]$(if ($Edited) { 7 } else { $profile.Gap })
        $frame.MarginLeft = [double]$insets[0]
        $frame.MarginTop = [double]$insets[1]
        $frame.MarginRight = [double]$insets[2]
        $frame.MarginBottom = [double]$insets[3]
        $frame.VerticalTextAlignment = if ($Edited) { 2 } else { $profile.Vertical }
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName frameLayout -NotePropertyValue ([pscustomobject]@{
            columns = [int]$frame.Columns
            columnSpacing = [double]$frame.ColumnSpacing
            left = [double]$frame.MarginLeft
            top = [double]$frame.MarginTop
            right = [double]$frame.MarginRight
            bottom = [double]$frame.MarginBottom
            verticalAlignment = [int]$frame.VerticalTextAlignment
        })
    }
    return 2
}
