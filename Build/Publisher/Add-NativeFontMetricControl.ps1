# Independent Latin feature-selection, fractional-size and line-advance controls.
function Add-NativeFontMetricControl {
    param($Publication, $FirstPage)

    $owner = $FirstPage
    $pages = 1
    $index = 0
    foreach ($family in @('Arial', 'Calibri', 'Times New Roman', 'Comic Sans MS', 'Segoe UI',
            'Lucida Console', 'Courier New', 'Consolas', 'Tahoma', 'Lucida Sans Unicode')) {
        $sizes = if ($family -eq 'Courier New') { @(12, 18, 12.5, 18.25) } else { @(12, 18) }
        $modes = if ($family -in @('Arial', 'Calibri', 'Times New Roman')) { @('Inherited', 'Single', 'Multiple') }
            elseif ($family -eq 'Courier New') { @('Single') } else { @('Single', 'Multiple') }
        foreach ($size in $sizes) {
            foreach ($mode in $modes) {
                if ($index -gt 0 -and $index % 3 -eq 0) {
                    $owner = $Publication.Pages.Add(1, $pages)
                    $objects.Add($owner)
                    $pages++
                }
                $name = 'FM' + $index + '-' + $size.ToString([Globalization.CultureInfo]::InvariantCulture) + '-' + $mode
                $shape = Add-TextFrame $owner $name ($name + " gyjp`rSecond line gyjp") 54 (54 + ($index % 3) * 228) 504 180 $size
                $frame = $shape.TextFrame
                $range = $frame.TextRange
                $paragraph = $range.ParagraphFormat
                $font = $range.Font
                $objects.Add($frame); $objects.Add($range); $objects.Add($paragraph); $objects.Add($font)
                # Set the complete range after text assignment; the default
                # typing font can otherwise leave a mixed-size terminal CR.
                $font.Name = $family
                $font.Size = [double]$size
                $paragraph.SpaceBefore = 0
                $paragraph.SpaceAfter = 0
                if ($mode -eq 'Single') { $paragraph.SetLineSpacing(0, 1) }
                if ($mode -eq 'Multiple') { $paragraph.SetLineSpacing(5, 1.19) }
                if ($font.Name -ne $family -or [Math]::Abs([double]$font.Size - $size) -gt .01) {
                    throw ('Native font mismatch for ' + $name)
                }
                $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName fontMetricsProfile -NotePropertyValue ([pscustomobject]@{
                    family = $font.Name
                    size = [double]$font.Size
                    rule = [int]$paragraph.LineSpacingRule
                    spacing = [double]$paragraph.LineSpacing
                    text = $range.Text
                })
                $index++
            }
        }
    }
    return $pages
}
