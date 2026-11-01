# Independent selected-design-height versus device-ascent alignment controls.
function Add-NativeFallbackAlignmentControl {
    param($Publication, $FirstPage)
    $owner = $FirstPage
    $pages = 1
    $index = 0
    foreach ($family in @('Courier New', 'Lucida Sans Unicode')) {
        foreach ($size in @(12, 18)) {
            foreach ($alignment in 0..2) {
                if ($index -gt 0 -and $index % 3 -eq 0) {
                    $owner = $Publication.Pages.Add(1, $pages)
                    $objects.Add($owner)
                    $pages++
                }
                $name = 'B' + $index + '-' + $family.Replace(' ', '') + '-' + $size + '-A' + $alignment
                $shape = Add-TextFrame $owner $name ($name + ' gyjp') 54 (54 + ($index % 3) * 228) 504 180 $size
                $frame = $shape.TextFrame
                $range = $frame.TextRange
                $font = $range.Font
                $paragraph = $range.ParagraphFormat
                $objects.Add($frame); $objects.Add($range); $objects.Add($font); $objects.Add($paragraph)
                $font.Name = $family
                $font.Size = [double]$size
                $frame.VerticalTextAlignment = $alignment
                $paragraph.SpaceBefore = 0
                $paragraph.SpaceAfter = 0
                $paragraph.SetLineSpacing(5, 1.19)
                if ($font.Name -ne $family -or [Math]::Abs([double]$font.Size - $size) -gt .01 -or [int]$frame.VerticalTextAlignment -ne $alignment) { throw 'Native fallback alignment mismatch.' }
                $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName fallbackAlignment -NotePropertyValue ([pscustomobject]@{
                    family = $font.Name
                    size = [double]$font.Size
                    alignment = [int]$frame.VerticalTextAlignment
                    factor = [double]$paragraph.LineSpacing
                    text = $range.Text
                })
                $index++
            }
        }
    }
    return $pages
}
