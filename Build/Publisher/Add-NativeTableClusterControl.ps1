function Add-NativeTableClusterControl {
    <#
    .SYNOPSIS
    Authors a dense square table whose cell and edge utilities approach or cross a native SPID cluster.
    .PARAMETER FirstPage
    The task-owned new publication page.
    .PARAMETER Size
    Eighteen fits within one utility cluster; twenty crosses its 1024-slot boundary.
    .PARAMETER Edited
    Fills the first normally invisible cell in green without changing the grid or borders.
    .OUTPUTS
    The physical page count.
    #>
    param($FirstPage, [ValidateSet(18,20)] [int] $Size, [switch] $Edited)

    $label = Add-TextFrame $FirstPage 'ClusterLabel' ('Independent dense {0}x{0} table' -f $Size) 36 60 480 42 12
    $label.TextFrame.ColumnSpacing = 0
    $shape = $FirstPage.Shapes.AddTable($Size, $Size, 36, 144, 480, 360)
    $shape.Name = 'ClusterTable'
    $table = $shape.Table
    $objects.Add($shape)
    $objects.Add($table)
    $table.ApplyAutoFormat(-2, $false, $false, $true, $true)
    for ($column = 1; $column -le $Size; $column++) { $table.Columns.Item($column).Width = 480 / $Size }
    for ($row = 1; $row -le $Size; $row++) { $table.Rows.Item($row).Height = 360 / $Size }
    for ($row = 1; $row -le $Size; $row++) {
        Set-NativePhase ('AuthorClusterTextRow-' + $row)
        for ($column = 1; $column -le $Size; $column++) {
            $range = $table.Cells($row,$column,$row,$column)
            $cell = $range.Item(1)
            $text = $cell.TextRange
            foreach ($value in @($range,$cell,$text)) { $objects.Add($value) }
            $cell.MarginLeft = 0; $cell.MarginRight = 0; $cell.MarginTop = 0; $cell.MarginBottom = 0
            $value = '{0:D2}{1:D2}' -f $row,$column
            $text.Text = $value
            $text.Font.Name = 'Arial'; $text.Font.Size = 8; $text.Font.Color.RGB = 0
            $text.Font.Bold = if ($row -eq 1) { -1 } else { 0 }
            $paragraph = $text.ParagraphFormat
            $objects.Add($paragraph)
            $paragraph.SpaceBefore = 0; $paragraph.SpaceAfter = 0; $paragraph.SetLineSpacing(4,12)
            $expectedText.Add($value)
        }
    }
    for ($row = 1; $row -le $Size; $row++) {
        Set-NativePhase ('AuthorClusterBorderRow-' + $row)
        for ($column = 1; $column -le $Size; $column++) {
            $range = $table.Cells($row,$column,$row,$column);$cell = $range.Item(1)
            $objects.Add($range);$objects.Add($cell)
            $sides = @('Left','Top')
            if ($column -eq $Size) { $sides += 'Right' }
            if ($row -eq $Size) { $sides += 'Bottom' }
            foreach ($side in $sides) {
                $border = $cell.('Border'+$side);$color = $border.Color
                $objects.Add($border);$objects.Add($color)
                $color.RGB = 0x008000 # RGB(0,128,0)
                $border.Weight = 0.5
            }
        }
    }
    # Border authoring can activate automatic white fills; set the desired
    # final cell paint only after all shared edges have been authored.
    for ($row = 1; $row -le $Size; $row++) {
        Set-NativePhase ('AuthorClusterFillRow-' + $row)
        for ($column = 1; $column -le $Size; $column++) {
            $range = $table.Cells($row,$column,$row,$column);$cell = $range.Item(1);$fill = $cell.Fill
            foreach ($value in @($range,$cell,$fill)) { $objects.Add($value) }
            $pattern = ($row+$column-2)%3
            if ($pattern -eq 2 -and -not($Edited -and $row -eq 1 -and $column -eq 3)) { $fill.Visible = 0 }
            else {
                $fill.Solid();$fill.Visible = -1;$fill.Transparency = 0
                $color = $fill.ForeColor;$objects.Add($color)
                $color.RGB = if($Edited -and $row -eq 1 -and $column -eq 3){0x00B4F0B4}elseif($pattern -eq 0){0x00A0E0FF}else{0x00FFE0A0}
            }
        }
    }
    $expectedShapes.Add([pscustomobject]@{name='ClusterTable';kind='Table';left=36;top=144;width=480;height=360;rows=$Size;columns=$Size})
    return 1
}
