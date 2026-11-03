function Add-NativeCandidateClusterTable {
    <#
    .SYNOPSIS
    Adds a fixed table to the isolated writable dense-table snapshot.
    .DESCRIPTION
    Requires one physical page with exactly one unmerged eighteen-by-eighteen
    table. Only the new ClusterRolloverProbe table is edited. Its eight fills
    and twenty-two unit edges require thirty table utility shapes, exceeding
    the thirteen slots remaining in the qualified near-boundary control.
    .PARAMETER Publication
    The task-owned writable snapshot opened by the native validation worker.
    .PARAMETER Objects
    Retains COM wrappers for the worker's lifetime.
    .OUTPUTS
    Probe identity, grid dimensions, row-major expected texts and retained table name.
    #>
    param(
        $Publication,
        [Collections.Generic.List[object]] $Objects
    )

    $pages = $Publication.Pages
    $Objects.Add($pages)
    if ($Publication.ReadOnly -or $pages.Count -ne 1) {
        throw 'Native cluster-table addition requires one writable physical page.'
    }
    $page = $pages.Item(1)
    $shapes = $page.Shapes
    $Objects.Add($page)
    $Objects.Add($shapes)
    $tables = [Collections.Generic.List[object]]::new()
    for ($index = 1; $index -le $shapes.Count; $index++) {
        $shape = $shapes.Item($index)
        $Objects.Add($shape)
        if ($shape.Name -ceq 'ClusterRolloverProbe') {
            throw 'The native cluster-table probe name already exists.'
        }
        if ($shape.Type -eq 18) {
            $tables.Add($shape)
        }
    }
    if ($tables.Count -ne 1) {
        throw 'Native cluster-table addition requires exactly one physical table.'
    }
    $retainedShape = $tables[0]
    $retainedTable = $retainedShape.Table
    $retainedRows = $retainedTable.Rows
    $retainedColumns = $retainedTable.Columns
    $Objects.Add($retainedTable)
    $Objects.Add($retainedRows)
    $Objects.Add($retainedColumns)
    if ($retainedRows.Count -ne 18 -or $retainedColumns.Count -ne 18) {
        throw 'Native cluster-table addition requires the eighteen-by-eighteen control table.'
    }
    for ($row = 1; $row -le 18; $row++) {
        Set-NativePhase ('InspectNativeClusterTableRow-' + $row)
        for ($column = 1; $column -le 18; $column++) {
            $range = $retainedTable.Cells($row, $column, $row, $column)
            $cell = $range.Item(1)
            $Objects.Add($range)
            $Objects.Add($cell)
            if ($cell.Width -ne 1 -or $cell.Height -ne 1) {
                throw 'Native cluster-table addition requires every retained cell to be unmerged.'
            }
        }
    }

    Set-NativePhase 'AddNativeClusterProbeTable'
    $shape = $shapes.AddTable(2, 4, 36, 540, 480, 48)
    $Objects.Add($shape)
    $shape.Name = 'ClusterRolloverProbe'
    $table = $shape.Table
    $rows = $table.Rows
    $columns = $table.Columns
    $Objects.Add($table)
    $Objects.Add($rows)
    $Objects.Add($columns)
    [void]$table.ApplyAutoFormat(-2, $false, $false, $true, $true) # pbTableAutoFormatNone
    for ($column = 1; $column -le 4; $column++) {
        $track = $columns.Item($column)
        $Objects.Add($track)
        $track.Width = 120
    }
    for ($row = 1; $row -le 2; $row++) {
        $track = $rows.Item($row)
        $Objects.Add($track)
        $track.Height = 24
    }
    $texts = [Collections.Generic.List[string]]::new()
    for ($row = 1; $row -le 2; $row++) {
        Set-NativePhase ('AuthorNativeClusterProbeRow-' + $row)
        for ($column = 1; $column -le 4; $column++) {
            $range = $table.Cells($row, $column, $row, $column)
            $cell = $range.Item(1)
            $text = $cell.TextRange
            $Objects.Add($range)
            $Objects.Add($cell)
            $Objects.Add($text)
            $cell.MarginLeft = 0
            $cell.MarginRight = 0
            $cell.MarginTop = 0
            $cell.MarginBottom = 0
            $value = 'C{0}{1}' -f $row, $column
            $text.Text = $value
            $texts.Add($value)
            $text = $cell.TextRange
            $font = $text.Font
            $color = $font.Color
            $paragraph = $text.ParagraphFormat
            $Objects.Add($text)
            $Objects.Add($font)
            $Objects.Add($color)
            $Objects.Add($paragraph)
            $font.Name = 'Arial'
            $font.Size = 8
            $font.Bold = 0
            $color.RGB = 0
            $paragraph.SpaceBefore = 0
            $paragraph.SpaceAfter = 0
            [void]$paragraph.SetLineSpacing(4, 12)
            $sides = @('Left', 'Top')
            if ($column -eq 4) {
                $sides += 'Right'
            }
            if ($row -eq 2) {
                $sides += 'Bottom'
            }
            foreach ($side in $sides) {
                $border = $cell.('Border' + $side)
                $color = $border.Color
                $Objects.Add($border)
                $Objects.Add($color)
                $color.RGB = 0x008000
                $border.Weight = 0.5
            }
        }
    }
    # Set final fills after border authoring, which can activate automatic white paint.
    for ($row = 1; $row -le 2; $row++) {
        Set-NativePhase ('FillNativeClusterProbeRow-' + $row)
        for ($column = 1; $column -le 4; $column++) {
            $range = $table.Cells($row, $column, $row, $column)
            $cell = $range.Item(1)
            $fill = $cell.Fill
            $color = $fill.ForeColor
            $Objects.Add($range)
            $Objects.Add($cell)
            $Objects.Add($fill)
            $Objects.Add($color)
            [void]$fill.Solid()
            $fill.Visible = -1
            $fill.Transparency = 0
            $color.RGB = 0x00B4F0B4
        }
    }
    return [ordered]@{
        shapeName = 'ClusterRolloverProbe'
        rows = 2
        columns = 4
        expectedTexts = $texts.ToArray()
        expectedUtilityShapes = 30
        retainedTableName = [string]$retainedShape.Name
    }
}
