# Synthetic whole-frame layout controls; native application access stays in the bounded corpus worker.
function Add-NativeFrameLayoutControl {
    param($Publication, $FirstPage)
    $profiles = @(
        @{ Name = 'Baseline'; Columns = 1; Gap = 0; Insets = @(0, 0, 0, 0); Vertical = 0 },
        @{ Name = 'Insets'; Columns = 1; Gap = 0; Insets = @(12, 6, 18, 9); Vertical = 0 },
        @{ Name = 'TwoColumns'; Columns = 2; Gap = 6; Insets = @(0, 0, 0, 0); Vertical = 0 },
        @{ Name = 'ThreeColumns'; Columns = 3; Gap = 9; Insets = @(0, 0, 0, 0); Vertical = 0 },
        @{ Name = 'Center'; Columns = 1; Gap = 0; Insets = @(0, 0, 0, 0); Vertical = 1; Short = $true },
        @{ Name = 'Bottom'; Columns = 1; Gap = 0; Insets = @(0, 0, 0, 0); Vertical = 2; Short = $true },
        @{ Name = 'Combined'; Columns = 2; Gap = 6; Insets = @(12, 6, 18, 9); Vertical = 1 },
        @{ Name = 'MixedRectangle'; Columns = 2; Gap = 6; Insets = @(9, 6, 9, 6); Vertical = 2; Short = $true; Mixed = $true },
        @{ Name = 'Fractional'; Columns = 3; Gap = 6.05; Insets = @(1.21, 2.42, 3.63, 4.84); Vertical = 0 }
    )
    $owner = $FirstPage
    $pages = 1
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0 -and $index % 3 -eq 0) {
            $owner = $Publication.Pages.Add(1, $pages)
            $objects.Add($owner)
            $pages++
        }
        $text = 'Frame layout control: ' + $profile.Name + "`r"
        $text += if ($profile.ContainsKey('Short')) { 'Short native body.' } else {
            (@(1..8 | ForEach-Object { "Paragraph $_ moves through the current columns and insets while the complete native story remains available." }) -join "`r")
        }
        $left = 54; $top = 54 + ($index % 3) * 228
        if ($profile.ContainsKey('Mixed')) {
            $shape = $owner.Shapes.AddShape(1, $left, $top, 504, 180)
            $shape.Name = $profile.Name
            $shape.Fill.ForeColor.RGB = 13434879
            $shape.Line.ForeColor.RGB = 0
            $objects.Add($shape)
            $frame = $shape.TextFrame
            $range = $frame.TextRange
            $objects.Add($frame); $objects.Add($range)
            $range.Text = $text; $range.Font.Name = 'Arial'; $range.Font.Size = 12; $range.Font.Color.RGB = 0
            $expectedText.Add($text)
            $expectedShapes.Add([pscustomobject]@{ name = $profile.Name; kind = 'TextRectangle'; left = $left; top = $top; width = 504; height = 180; font = 'Arial'; size = 12 })
        } else {
            $shape = Add-TextFrame $owner $profile.Name $text $left $top 504 180
            $frame = $shape.TextFrame
            $objects.Add($frame)
        }
        $frame.Columns = $profile.Columns
        $frame.ColumnSpacing = [double]$profile.Gap
        $frame.MarginLeft = [double]$profile.Insets[0]; $frame.MarginTop = [double]$profile.Insets[1]
        $frame.MarginRight = [double]$profile.Insets[2]; $frame.MarginBottom = [double]$profile.Insets[3]
        $frame.VerticalTextAlignment = $profile.Vertical
        $attributes = [pscustomobject]@{
            columns = [int]$frame.Columns; columnSpacing = [double]$frame.ColumnSpacing
            left = [double]$frame.MarginLeft; top = [double]$frame.MarginTop
            right = [double]$frame.MarginRight; bottom = [double]$frame.MarginBottom
            verticalAlignment = [int]$frame.VerticalTextAlignment
        }
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName frameLayout -NotePropertyValue $attributes
    }
    return $pages
}
