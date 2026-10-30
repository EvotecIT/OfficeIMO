# Synthetic whole-frame controls. No selection, Characters access or user file is used.
function Add-NativeCharacterAttributeControl {
    param($Publication, $FirstPage, [switch] $UnderlinePatterns)
    $profiles = @(
        @{ Name = 'Regular' },
        @{ Name = 'Italic'; Italic = $true },
        @{ Name = 'UnderlineSingle'; Underline = 1 },
        @{ Name = 'UnderlineDouble'; Underline = 3 },
        @{ Name = 'Strikethrough'; StrikeThrough = $true },
        @{ Name = 'Superscript'; SuperScript = $true },
        @{ Name = 'Subscript'; SubScript = $true },
        @{ Name = 'FontCambria'; Family = 'Cambria' },
        @{ Name = 'FontCourier'; Family = 'Courier New' },
        @{ Name = 'FontConsolas'; Family = 'Consolas' },
        @{ Name = 'ColorRed'; Rgb = 255 },
        @{ Name = 'Combined'; Family = 'Courier New'; Size = 14; Bold = $true; Italic = $true; Underline = 3; StrikeThrough = $true; Rgb = 16711680 }
    )
    if ($UnderlinePatterns) {
        $profiles = @(0..17 | ForEach-Object { @{ Name = 'Underline' + $_; Underline = $_ } })
    }
    $owner = $FirstPage
    $pages = 1
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0 -and $index % 6 -eq 0) {
            $owner = $Publication.Pages.Add(1, $pages)
            $objects.Add($owner)
            $pages++
        }
        $size = if ($profile.ContainsKey('Size')) { $profile.Size } else { 12 }
        $family = if ($profile.ContainsKey('Family')) { $profile.Family } else { 'Arial' }
        $rgb = if ($profile.ContainsKey('Rgb')) { $profile.Rgb } else { 0 }
        $underline = if ($profile.ContainsKey('Underline')) { $profile.Underline } else { 0 }
        $shape = Add-TextFrame $owner $profile.Name ('Native character control: ' + $profile.Name) 54 (54 + ($index % 6) * 108) 504 72 $size
        $font = $shape.TextFrame.TextRange.Font
        $objects.Add($font)
        $font.Name = $family
        $font.Underline = $underline
        $font.Color.RGB = $rgb
        $attributes = [ordered]@{ family = $family; size = $size; rgb = $rgb; underline = $underline }
        foreach ($property in @('Bold', 'Italic', 'StrikeThrough', 'SuperScript', 'SubScript')) {
            $enabled = $profile.ContainsKey($property) -and $profile[$property]
            $font.$property = $(if ($enabled) { -1 } else { 0 })
            $attributes[$property] = [bool]$enabled
        }
        $expectedShapes[$expectedShapes.Count - 1].font = $family
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName characterAttributes -NotePropertyValue ([pscustomobject]$attributes)
    }
    return $pages
}
