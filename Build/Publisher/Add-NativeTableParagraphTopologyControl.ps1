# First and final cell controls share fitting geometry and unchanged neighbors.
function Add-NativeTableParagraphTopologyControl {
    param($Publication, $FirstPage, [switch] $EmptyParagraphs, [switch] $UniformEmpty, [switch] $TextReplacements)
    $profiles = @(Get-NativeParagraphTopologyProfiles -EmptyParagraphs:$EmptyParagraphs -UniformEmpty:$UniformEmpty -TextReplacements:$TextReplacements)
    $owner = $FirstPage
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0) {
            $owner = $Publication.Pages.Add(1, $index)
            $objects.Add($owner)
        }
        Set-NativePhase ('TableParagraphTopology-' + $profile.Name)
        $table = $owner.Shapes.AddTable(2, 2, 72, 72, 468, 432)
        $table.Name = $profile.Name
        $objects.Add($table)
        $values = @($profile.Text, 'Neighbor one', 'Neighbor two', $profile.Text)
        $observed = [Collections.Generic.List[object]]::new()
        for ($cellIndex = 0; $cellIndex -lt 4; $cellIndex++) {
            $row = [int][Math]::Floor($cellIndex / 2) + 1
            $column = $cellIndex % 2 + 1
            $cell = $table.Table.Cells($row, $column, $row, $column).Item(1)
            $objects.Add($cell)
            $cell.TextRange.Text = $values[$cellIndex]
            $cell.TextRange.Font.Name = 'Arial'
            $cell.TextRange.Font.Size = 12
            $cell.TextRange.Font.Color.RGB = 0
            $cell.TextRange.Font.Bold = 0
            $cell.Fill.ForeColor.RGB = 16777215
            if ($cellIndex -in @(0, 3)) {
                $whole = $cell.TextRange
                $objects.Add($whole)
                $paragraphs = @(Set-NativeParagraphTopologyFormatting -WholeRange $whole -Profile $profile)
                $observed.Add([pscustomobject]@{ cell = $cellIndex; order = $profile.Order; observed = $paragraphs })
            }
        }
        $expectedShapes.Add([pscustomobject]@{
            name = $profile.Name
            kind = 'Table'
            rows = 2
            columns = 2
            cells = $values
            left = 72
            top = 72
            width = 468
            height = 432
            paragraphTopology = $observed.ToArray()
        })
        foreach ($value in $values) { $expectedText.Add($value) }
    }
    return $profiles.Count
}
