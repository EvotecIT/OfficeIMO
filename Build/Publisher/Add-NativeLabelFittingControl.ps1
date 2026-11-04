# Independent labelled tight-height and linked continuation controls.
function Add-NativeLabelFittingControl {
    param($Publication, $FirstPage)

    $profiles = [Collections.Generic.List[object]]::new()
    foreach ($labelSize in @(12, 24)) {
        foreach ($rule in @(4, 5)) {
            foreach ($height in @(6, 12, 18, 35, 36)) {
                $profiles.Add(@{ Kind = 'Tight'; LabelSize = $labelSize; Rule = $rule; Height = $height; Width = 300 })
            }
        }
    }
    foreach ($height in @(12, 36, 12)) {
        $profiles.Add(@{ Kind = 'Linked'; LabelSize = 12; Rule = 4; Height = $height; Width = 110 })
    }
    $linked = [Collections.Generic.List[object]]::new()
    $owner = $FirstPage
    $pages = 1
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0 -and $index % 3 -eq 0) {
            $owner = $Publication.Pages.Add(1, $pages)
            $objects.Add($owner)
            $pages++
        }
        $name = 'LF' + $index
        $text = if ($profile.Kind -eq 'Linked') {
            if ($linked.Count -eq 0) { @(0..29 | ForEach-Object { 'T' + $_.ToString('D2') }) -join ' ' } else { '' }
        } else { "First gyjp`rSecond gyjp`rThird gyjp" }
        # Apply the complete source's formatting before its frame is narrowed.
        $shape = Add-TextFrame $owner $name $text 54 (54 + ($index % 3) * 228) 504 180 12
        $frame = $shape.TextFrame
        $range = $frame.TextRange
        $font = $range.Font
        $paragraph = $range.ParagraphFormat
        $objects.Add($frame); $objects.Add($range); $objects.Add($font); $objects.Add($paragraph)
        $range.LanguageID = 1033
        $font.Name = 'Arial'; $font.Size = 12
        $paragraph.SpaceBefore = 0; $paragraph.SpaceAfter = 0
        $paragraph.KeepLinesTogether = 0; $paragraph.KeepWithNext = 0; $paragraph.WidowControl = 0
        # Applying a list to an empty target creates a paragraph and prevents
        # linking. The source's list formatting flows into linked empty targets.
        if ($text.Length -gt 0) {
            $paragraph.SetListType(23, '*')
            $paragraph.ListBulletFontName = 'Arial'
            $paragraph.ListBulletFontSize = $profile.LabelSize
            $paragraph.ListIndent = 18
        }
        $spacing = if ($profile.Rule -eq 4) { 18 } else { 1.19 }
        $paragraph.SetLineSpacing($profile.Rule, [double]$spacing)
        $frame.VerticalTextAlignment = 0
        $settings = [pscustomobject]@{
            kind = $profile.Kind; sourceText = $text; family = [string]$font.Name; size = [double]$font.Size
            languageId = [int]$range.LanguageID; listType = [int]$paragraph.ListType
            marker = [string]$paragraph.ListBulletText; labelFamily = [string]$paragraph.ListBulletFontName
            labelSize = [double]$paragraph.ListBulletFontSize; listIndent = [double]$paragraph.ListIndent
            leftIndent = [double]$paragraph.LeftIndent; firstLineIndent = [double]$paragraph.FirstLineIndent
            rule = [int]$paragraph.LineSpacingRule; spacing = [double]$paragraph.LineSpacing
            frameWidth = $profile.Width; frameHeight = $profile.Height
        }
        if ($text.Length -gt 0 -and ($settings.listType -ne 23 -or $settings.marker -ne '*' -or $settings.labelFamily -ne 'Arial' -or
            [Math]::Abs($settings.labelSize - $profile.LabelSize) -gt .01)) {
            throw ('Native list settings differ for ' + $name)
        }
        $shape.Width = [double]$profile.Width; $shape.Height = [double]$profile.Height
        $wrap = $shape.TextWrap; $objects.Add($wrap); $wrap.Type = 0
        if ($wrap.Type -ne 0) { throw 'Label fitting controls require object exclusions disabled.' }
        $expectedShapes[$expectedShapes.Count - 1].width = $profile.Width
        $expectedShapes[$expectedShapes.Count - 1].height = $profile.Height
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName labelFittingProfile -NotePropertyValue $settings
        if ($profile.Kind -eq 'Linked') { $linked.Add($shape) }
    }
    # Physical placement differs from source reading order.
    $linked[0].TextFrame.NextLinkedTextFrame = $linked[2].TextFrame
    $linked[2].TextFrame.NextLinkedTextFrame = $linked[1].TextFrame
    return $pages
}
