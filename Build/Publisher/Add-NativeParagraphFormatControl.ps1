# Synthetic whole-frame controls; no selection or character-range probing.
function Add-NativeParagraphFormatControl {
    param($Publication, $FirstPage, [switch] $AlignedTabs)
    $profiles = @(
        @{ Name = 'Default' },
        @{ Name = 'Center'; Alignment = 1 },
        @{ Name = 'Right'; Alignment = 2 },
        @{ Name = 'Justify'; Alignment = 6 },
        @{ Name = 'Spacing'; SpaceBefore = 6; SpaceAfter = 9 },
        @{ Name = 'Insets'; LeftIndent = 36; RightIndent = 24 },
        @{ Name = 'FirstLine'; FirstLineIndent = 18 },
        @{ Name = 'Hanging'; LeftIndent = 36; FirstLineIndent = -18 },
        @{ Name = 'Exactly18'; Rule = 4; Spacing = 18 },
        @{ Name = 'Multiple1pt5'; Rule = 5; Spacing = 1.5 },
        @{ Name = 'Single'; Rule = 0 },
        @{ Name = 'Double'; Rule = 2 },
        @{ Name = 'Tabs'; Tabs = @(72, 144) },
        @{ Name = 'Combined'; Alignment = 1; LeftIndent = 36; RightIndent = 18; FirstLineIndent = 12; SpaceBefore = 6; SpaceAfter = 9; Rule = 5; Spacing = 1.25; Tabs = @(72, 144) }
    )
    if ($AlignedTabs) {
        $profiles = @(
            @{ Name = 'LeftTabs'; Alignment = 0; Tabs = @(72, 144) },
            @{ Name = 'CenterTabs'; Alignment = 1; Tabs = @(72, 144) },
            @{ Name = 'RightTabs'; Alignment = 2; Tabs = @(72, 144) }
        )
    }
    $owner = $FirstPage
    $pages = 1
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0 -and $index % 4 -eq 0) {
            $owner = $Publication.Pages.Add(1, $pages)
            $objects.Add($owner)
            $pages++
        }
        $text = 'Native paragraph control: ' + $profile.Name + '. This sentence wraps so paragraph spacing and indentation are visible.' + "`rSecond paragraph with the same native settings."
        if ($profile.ContainsKey('Tabs')) {
            $text = 'Native paragraph control: ' + $profile.Name + "`rLeft`tMiddle`tRight"
        }
        $shape = Add-TextFrame $owner $profile.Name $text 54 (54 + ($index % 4) * 168) 504 144
        $paragraph = $shape.TextFrame.TextRange.ParagraphFormat
        $objects.Add($paragraph)
        $attributes = [ordered]@{}
        foreach ($property in @('Alignment', 'LeftIndent', 'RightIndent', 'FirstLineIndent', 'SpaceBefore', 'SpaceAfter')) {
            $value = if ($profile.ContainsKey($property)) { $profile[$property] } else { 0 }
            $paragraph.$property = $value
            $attributes[$property] = $paragraph.$property
        }
        $rule = if ($profile.ContainsKey('Rule')) { $profile.Rule } else { 0 }
        $spacing = if ($profile.ContainsKey('Spacing')) { $profile.Spacing } else { 1 }
        $paragraph.SetLineSpacing($rule, $spacing)
        $attributes.LineSpacingRule = $paragraph.LineSpacingRule
        $attributes.LineSpacing = $paragraph.LineSpacing
        $tabs = $paragraph.Tabs
        $objects.Add($tabs)
        $tabs.ClearAll()
        $positions = @()
        if ($profile.ContainsKey('Tabs')) {
            $positions = $profile.Tabs
            foreach ($position in $positions) {
                $stop = $tabs.Add($position, 0, 0)
                $objects.Add($stop)
            }
        }
        $attributes.TabPositions = $positions
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName paragraphAttributes -NotePropertyValue ([pscustomobject]$attributes)
    }
    return $pages
}
