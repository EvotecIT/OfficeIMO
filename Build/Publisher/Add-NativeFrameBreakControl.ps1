# Independent exact-line placement, tight-height and short-token continuation controls.
function Add-NativeFrameBreakControl {
    param($Publication, $FirstPage)

    $profiles = [Collections.Generic.List[object]]::new()
    # One frame at the same page origin separates metric rounding from placement.
    foreach ($family in @('Arial', 'Courier New')) {
        foreach ($line in @(9.6, 10)) {
            $profiles.Add(@{ Kind = 'CompressedMetric'; Family = $family; Line = $line; Alignment = 0; Height = 180; Width = 300; OwnPage = $true })
        }
    }
    foreach ($height in @(12, 14, 15.4, 15.6, 16, 16.4)) {
        $profiles.Add(@{ Kind = 'FirstExact'; Family = 'Arial'; Line = 18; Alignment = 0; Height = $height; Width = 300; OwnPage = $true })
    }
    foreach ($height in @(8, 10, 10.6, 11.5, 12, 12.5)) {
        $profiles.Add(@{ Kind = 'FirstMultiple'; Family = 'Arial'; Multiple = 1.19; Alignment = 0; Height = $height; Width = 300; OwnPage = $true })
    }
    foreach ($family in @('Arial', 'Courier New')) {
        $spacing = if ($family -eq 'Arial') { @(8, 12, 18, 24) } else { @(12, 18) }
        foreach ($line in $spacing) {
            foreach ($alignment in 0..2) {
                $profiles.Add(@{ Kind = 'Placement'; Family = $family; Line = $line; Alignment = $alignment; Height = 180; Width = 300 })
            }
        }
    }
    foreach ($height in @(17, 18, 35, 36, 53, 54)) {
        $profiles.Add(@{ Kind = 'TightExact'; Family = 'Arial'; Line = 18; Alignment = 0; Height = $height; Width = 300 })
    }
    foreach ($height in @(12.9, 13.3, 28.4, 28.8, 44, 44.4)) {
        $profiles.Add(@{ Kind = 'TightMultiple'; Family = 'Arial'; Multiple = 1.19; Alignment = 0; Height = $height; Width = 300 })
    }
    foreach ($height in @(35, 36)) {
        $profiles.Add(@{ Kind = 'Tokens'; Family = 'Arial'; Line = 18; Alignment = 0; Height = $height; Width = 100 })
    }
    foreach ($height in @(36, 72, 36)) {
        $profiles.Add(@{ Kind = 'Linked'; Family = 'Arial'; Line = 18; Alignment = 0; Height = $height; Width = 100 })
    }
    $tokens = (@(0..39 | ForEach-Object { 'T' + $_.ToString('D2') }) -join ' ')
    $linked = [Collections.Generic.List[object]]::new()
    $owner = $FirstPage
    $pages = 1
    $ordinary = 0
    $previousOwnPage = $false
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        $ownsPage = $profile.ContainsKey('OwnPage')
        if ($index -gt 0 -and ($ownsPage -or $previousOwnPage -or ($ordinary -gt 0 -and $ordinary % 3 -eq 0))) {
            $owner = $Publication.Pages.Add(1, $pages)
            $objects.Add($owner)
            $pages++
        }
        $name = 'FB' + $index
        $text = if ($profile.Kind -eq 'Linked' -and $linked.Count -gt 0) { '' }
            elseif ($profile.Kind -in @('Tokens', 'Linked')) { $tokens }
            else { $name + " gyjp`rSecond gyjp`rThird gyjp" }
        # Format the complete source in a roomy frame, then narrow it. A tight
        # frame's TextRange getter is only its assigned portion of the story.
        $top = if ($ownsPage) { 54 } else { 54 + ($ordinary % 3) * 228 }
        $shape = Add-TextFrame $owner $name $text 54 $top 504 180 12
        if (!$ownsPage) { $ordinary++ }
        $previousOwnPage = $ownsPage
        $frame = $shape.TextFrame
        $range = $frame.TextRange
        $font = $range.Font
        $paragraph = $range.ParagraphFormat
        $objects.Add($frame); $objects.Add($range); $objects.Add($font); $objects.Add($paragraph)
        $font.Name = $profile.Family
        $font.Size = [double]12
        $paragraph.SpaceBefore = 0
        $paragraph.SpaceAfter = 0
        if ($profile.ContainsKey('Multiple')) { $paragraph.SetLineSpacing(5, [double]$profile.Multiple) }
        else { $paragraph.SetLineSpacing(4, [double]$profile.Line) }
        $frame.VerticalTextAlignment = $profile.Alignment
        if ($font.Name -ne $profile.Family -or [Math]::Abs([double]$font.Size - 12) -gt .01) {
            throw ('Native frame-break font mismatch for ' + $name)
        }
        $shape.Width = [double]$profile.Width
        $shape.Height = [double]$profile.Height
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        $wrap.Type = 0
        if ($wrap.Type -ne 0) { throw 'Frame-break controls require wrapping exclusions disabled.' }
        $expectedShapes[$expectedShapes.Count - 1].width = $profile.Width
        $expectedShapes[$expectedShapes.Count - 1].height = $profile.Height
        $expectedShapes[$expectedShapes.Count - 1].font = $profile.Family
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName frameBreakProfile -NotePropertyValue ([pscustomobject]@{
            kind = $profile.Kind
            family = $profile.Family
            size = 12
            rule = [int]$paragraph.LineSpacingRule
            spacing = [double]$paragraph.LineSpacing
            alignment = $profile.Alignment
            frameWidth = [double]$shape.Width
            frameHeight = [double]$shape.Height
            sourceText = $text
        })
        if ($profile.Kind -eq 'Linked') { $linked.Add($shape) }
    }
    # Physical placement is deliberately distinct from the chain's reading order.
    $linked[0].TextFrame.NextLinkedTextFrame = $linked[2].TextFrame
    $linked[2].TextFrame.NextLinkedTextFrame = $linked[1].TextFrame
    return $pages
}
