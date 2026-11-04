# Native Words ranges author real run boundaries without a selection or clipboard.
function Set-NativeMixedRunFormatting {
    param($ParagraphRange, [bool] $FrenchMiddle, [hashtable] $Overrides = @{})
    $words = if ($Overrides.ContainsKey(2)) { @(1, 2, 3) } else { @(1, 2) }
    foreach ($word in $words) {
        $range = $ParagraphRange.Words($word, 1)
        $objects.Add($range)
        $font = $range.Font
        $objects.Add($font)
        if ($word -le 2) {
            $font.Size = $(if ($word -eq 1) { 14 } else { 12 })
            $font.Bold = $(if ($word -eq 1) { -1 } else { 0 })
            $font.Italic = $(if ($word -eq 2) { -1 } else { 0 })
            $font.Color.RGB = $(if ($word -eq 1) { 16711680 } else { 255 })
        }
        if ($Overrides.ContainsKey($word - 1)) {
            $override = $Overrides[$word - 1]
            foreach ($name in 'Size', 'Bold', 'Italic', 'Underline') {
                if ($override.ContainsKey($name)) { $font.$name = $override[$name] }
            }
            if ($override.ContainsKey('Family')) { $font.Name = $override.Family }
            if ($override.ContainsKey('Rgb')) { $font.Color.RGB = $override.Rgb }
        }
        if ($FrenchMiddle -and $word -eq 2) { $range.LanguageID = 1036 }
        [pscustomobject]@{
            word = $word
            text = [string]$range.Text
            size = [double]$font.Size
            bold = [int]$font.Bold
            italic = [int]$font.Italic
            rgb = [int]$font.Color.RGB
            languageId = [int]$range.LanguageID
        }
    }
}
