# Explicit style inheritance makes moved empty paragraphs independent of the frame's typing default.
function Add-NativeEmptyParagraphTypingStyles {
    param($Publication)
    $normal = $Publication.TextStyles.Item(1)
    $objects.Add($normal)
    $styles = $Publication.TextStyles
    $objects.Add($styles)
    # Supply the inherited objects explicitly: PS7 cannot marshal this method's
    # omitted/null optional COM objects through the installed Publisher interface.
    $baseFont = $normal.Font
    $baseParagraphs = $normal.ParagraphFormat
    $objects.Add($baseFont)
    $objects.Add($baseParagraphs)
    $result = @{ '3' = [string]$normal.Name }
    $observations = [Collections.Generic.List[object]]::new()
    foreach ($spec in @(@{ Size = 10 }, @{ Size = 12 }, @{ Size = 14; RtlEmphasis = $true }, @{ Size = 12; RtlSize = 9; WesternEmphasis = $true })) {
        $size = $spec.Size
        $key = 'size' + $size + $(if ($spec.ContainsKey('RtlSize')) { 'bi' + $spec.RtlSize })
        $typed = $styles.Add(('Retained empty Arial ' + $key), [string]$normal.Name, $baseFont, $baseParagraphs)
        $objects.Add($typed)
        $font = $typed.Font
        $objects.Add($font)
        $font.Name = 'Arial'
        $font.Size = $size
        if ($spec.ContainsKey('RtlSize')) { $font.SizeBi = $spec.RtlSize }
        $font.Bold = $(if ($spec.ContainsKey('WesternEmphasis')) { -1 } else { 0 })
        $font.Italic = $(if ($spec.ContainsKey('WesternEmphasis')) { -1 } else { 0 })
        $font.BoldBi = $(if ($spec.ContainsKey('RtlEmphasis')) { -1 } else { 0 })
        $font.ItalicBi = $(if ($spec.ContainsKey('RtlEmphasis')) { -1 } else { 0 })
        $font.Color.RGB = 0
        $typed.Font = $font
        $result[$key] = [string]$typed.Name
        $observations.Add([pscustomobject]@{ name = [string]$typed.Name; requestedSize = $size; observedSize = [double]$typed.Font.Size; observedFamily = [string]$typed.Font.Name; observedRtlSize = [double]$typed.Font.SizeBi; observedBold = [int]$typed.Font.Bold; observedItalic = [int]$typed.Font.Italic; observedRtlBold = [int]$typed.Font.BoldBi; observedRtlItalic = [int]$typed.Font.ItalicBi })
    }
    $result['1'] = $result['size12']
    $result.observations = $observations.ToArray()
    return $result
}
