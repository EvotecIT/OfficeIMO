# Native word breaks under explicit US English, plus a linked continuation.
function Add-NativeHyphenationControl {
    param($Publication, $FirstPage)

    $options = $application.Options
    $objects.Add($options)
    # Observe the existing defaults; validation never changes application options.
    $automatic = [bool]$options.AutoHyphenate
    $zone = [double]$options.HyphenationZone
    $profiles = [Collections.Generic.List[object]]::new()
    foreach ($word in @('creative', 'extraordinary')) {
        foreach ($width in @(28, 29, 30, 32, 36)) {
            $profiles.Add(@{ Text = 'A ' + $word; Width = $width; Height = 180; Kind = 'NarrowWord' })
        }
    }
    foreach ($width in @(100, 216)) {
        $profiles.Add(@{ Text = 'publication hyphenation automatic creative extraordinary preserve international'; Width = $width; Height = 180; Kind = 'Words' })
    }
    $story = ('Publication stories preserve reading order across publication pages. ' * 8).TrimEnd()
    foreach ($index in 0..2) {
        $profiles.Add(@{ Text = $(if ($index -eq 0) { $story } else { '' }); Width = 100; Height = 36; Kind = 'Linked' })
    }
    $linked = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        $page = if ($index -eq 0) { $FirstPage } else { $Publication.Pages.Add(1, $index) }
        if ($index -gt 0) { $objects.Add($page) }
        $shape = Add-TextFrame $page ('HY' + $index) $profile.Text 54 54 504 180 12
        $frame = $shape.TextFrame
        $range = $frame.TextRange
        $font = $range.Font
        $objects.Add($frame); $objects.Add($range); $objects.Add($font)
        $range.LanguageID = 1033
        $shape.Width = [double]$profile.Width
        $shape.Height = [double]$profile.Height
        $wrap = $shape.TextWrap
        $objects.Add($wrap)
        $wrap.Type = 0
        $expected = $expectedShapes[$expectedShapes.Count - 1]
        $expected.width = $profile.Width
        $expected.height = $profile.Height
        $expected | Add-Member -NotePropertyName hyphenationProfile -NotePropertyValue ([pscustomobject]@{
            kind = $profile.Kind; sourceText = $profile.Text; languageId = 1033
            applicationDefaultAutoHyphenate = $automatic; applicationHyphenationZonePoints = $zone
            automaticPairKerningThresholdPoints = [double]$font.AutomaticPairKerningThreshold
            manualKerningPoints = [double]$font.Kerning
        })
        if ($profile.Kind -eq 'Linked') { $linked.Add($shape) }
    }
    $linked[0].TextFrame.NextLinkedTextFrame = $linked[2].TextFrame
    $linked[2].TextFrame.NextLinkedTextFrame = $linked[1].TextFrame
    return $profiles.Count
}
