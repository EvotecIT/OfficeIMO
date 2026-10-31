# Independent final-paragraph spacing controls for native text alignment.
function Add-NativeVerticalPlacementControl {
    param($Publication, $FirstPage)

    $profiles = [Collections.Generic.List[object]]::new()
    foreach ($alignment in 0..2) {
        foreach ($terminal in @($false, $true)) {
            foreach ($paragraphs in 1..2) {
                $profiles.Add(@{ Alignment = $alignment; Terminal = $terminal; Paragraphs = $paragraphs; After = 6; Size = 18 })
            }
        }
    }
    foreach ($after in @(0, 12)) {
        foreach ($alignment in 1..2) {
            foreach ($paragraphs in 1..2) {
                $profiles.Add(@{ Alignment = $alignment; Terminal = $true; Paragraphs = $paragraphs; After = $after; Size = 18 })
            }
        }
    }
    foreach ($paragraphs in 1..2) {
        $profiles.Add(@{ Alignment = 2; Terminal = $true; Paragraphs = $paragraphs; After = 6; Size = 12 })
    }
    $owner = $FirstPage
    $pages = 1
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0 -and $index % 3 -eq 0) {
            $owner = $Publication.Pages.Add(1, $pages)
            $objects.Add($owner)
            $pages++
        }
        $name = 'V' + $index + '-A' + $profile.Alignment + '-P' + $profile.Paragraphs + '-S' + $profile.After + '-F' + $profile.Size + '-CR' + [int]$profile.Terminal
        $text = $name + ' gyjp'
        if ($profile.Paragraphs -eq 2) { $text += "`rSecond paragraph gyjp" }
        if ($profile.Terminal) { $text += "`r" }
        $shape = Add-TextFrame $owner $name $text 54 (54 + ($index % 3) * 228) 504 180 $profile.Size
        $frame = $shape.TextFrame
        $range = $frame.TextRange
        $paragraph = $range.ParagraphFormat
        $objects.Add($frame); $objects.Add($range); $objects.Add($paragraph)
        $frame.VerticalTextAlignment = $profile.Alignment
        $paragraph.SpaceBefore = 0
        $paragraph.SpaceAfter = $profile.After
        if ([int]$frame.VerticalTextAlignment -ne $profile.Alignment -or [double]$paragraph.SpaceAfter -ne $profile.After) {
            throw 'Native placement setter mismatch.'
        }
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName placementProfile -NotePropertyValue ([pscustomobject]@{
            alignment = $profile.Alignment; terminal = $profile.Terminal; paragraphs = $profile.Paragraphs
            spaceAfter = [double]$paragraph.SpaceAfter; fontSize = $profile.Size; observedText = $range.Text
        })
    }
    return $pages
}
