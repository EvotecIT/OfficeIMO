# Synthetic table/cell controls; native access stays in the bounded corpus worker.
function Add-NativeTableControl {
    param($Publication, $FirstPage, [string] $Case, [string] $PicturePath)
    if ($Case -eq 'TableLabel') {
        # Escapes retain PS5.1 syntax/encoding compatibility.
        $multilingual = 'Za' + [char]0x017C + [char]0x00F3 + [char]0x0142 + [char]0x0107 + ' g' + [char]0x0119 + [char]0x015B + 'l' + [char]0x0105 + ' ja' + [char]0x017A + [char]0x0144 + ' / ' + [char]0x0391 + [char]0x03BB + [char]0x03C6 + [char]0x03B1
        $null = Add-TextFrame $FirstPage 'LabelTitle' $multilingual 18 12 252 42 12
        $profiles = @(@{ Name = 'LabelTable' })
    } else {
        $profiles = @(
            @{ Name = 'Baseline' },
            @{ Name = 'ClearFirst'; Clear = @(0) },
            @{ Name = 'ClearFinal'; Clear = @(3) },
            @{ Name = 'ClearStyledFirst'; Clear = @(0); Styled = $true },
            @{ Name = 'FrenchFirst'; Language = 1036 },
            @{ Name = 'ClearFrenchFirst'; Language = 1036; Clear = @(0) },
            @{ Name = 'RefillFrenchStyledFirst'; Language = 1036; Styled = $true; Clear = @(0); Refill = $true },
            @{ Name = 'ClearThenStyleFirst'; Clear = @(0); StyleAfterClear = $true },
            @{ Name = 'ClearAll'; Clear = @(0, 1, 2, 3) }
        )
        if ($Case -eq 'TableCellTypingEarlyRange') { $profiles = @($profiles | Where-Object { !$_.ContainsKey('StyleAfterClear') }) }
    }
    $owner = $FirstPage; $pages = 1
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0 -and $index % 4 -eq 0) {
            $owner = $Publication.Pages.Add(1, $pages); $objects.Add($owner); $pages++
        }
        if ($Case -eq 'TableLabel') {
            $left = 18; $top = 66; $width = 180; $height = 72
        } else {
            $left = 72; $top = 64 + ($index % 4) * 176; $width = 360; $height = 104
            $null = Add-TextFrame $owner ('Caption-' + $profile.Name) $profile.Name $left ($top - 22) 432 20 12
        }
        Set-NativePhase ('Table-' + $profile.Name)
        $table = $owner.Shapes.AddTable(2, 2, $left, $top, $width, $height)
        $table.Name = $profile.Name; $objects.Add($table)
        $values = @('Item', 'Count', 'Alpha', '2')
        $observed = [Collections.Generic.List[object]]::new()
        for ($row = 1; $row -le 2; $row++) {
            for ($column = 1; $column -le 2; $column++) {
                $cellIndex = ($row - 1) * 2 + $column - 1
                $cell = $table.Table.Cells($row, $column, $row, $column).Item(1); $objects.Add($cell)
                # Preserve TableLabel's fresh-range calls. The early-range case
                # deliberately captures before Text assignment and can exclude CR.
                if ($Case -eq 'TableCellTypingEarlyRange') {
                    $range = $cell.TextRange; $objects.Add($range)
                    $range.Text = $values[$cellIndex]
                    $range.Font.Name = 'Arial'; $range.Font.Size = 10; $range.Font.Color.RGB = 0
                } else {
                    $cell.TextRange.Text = $values[$cellIndex]
                    $cell.TextRange.Font.Name = 'Arial'; $cell.TextRange.Font.Size = 10; $cell.TextRange.Font.Color.RGB = 0
                }
                if ($row -eq 1) {
                    $cell.Fill.ForeColor.RGB = 15790320
                    if ($Case -eq 'TableCellTypingEarlyRange') { $range.Font.Bold = -1 } else { $cell.TextRange.Font.Bold = -1 }
                }
                else { $cell.Fill.ForeColor.RGB = 16777215 }
                if ($Case -ne 'TableLabel') {
                    if ($Case -ne 'TableCellTypingEarlyRange') { $range = $cell.TextRange; $objects.Add($range) }
                    if ($cellIndex -eq 0 -and $profile.ContainsKey('Language')) { $range.LanguageID = [int]$profile.Language }
                    if ($cellIndex -eq 0 -and $profile.ContainsKey('Styled')) {
                        $range.Font.Name = 'Times New Roman'; $range.Font.Size = 14
                        $range.Font.Color.RGB = 16711680; $range.Font.Bold = -1; $range.Font.Italic = -1
                    }
                    $before = [string]$range.Text; $beforeLanguage = [int]$range.LanguageID
                    if ($profile.ContainsKey('Clear') -and $profile.Clear -contains $cellIndex) {
                        $range.Text = ''; $values[$cellIndex] = ''
                        if ($profile.ContainsKey('Refill')) { $range.Text = 'Refilled'; $values[$cellIndex] = 'Refilled' }
                        if ($profile.ContainsKey('StyleAfterClear')) {
                            $range = $cell.TextRange; $objects.Add($range)
                            $range.Font.Name = 'Times New Roman'; $range.Font.Size = 14
                            $range.Font.Color.RGB = 16711680; $range.Font.Bold = -1; $range.Font.Italic = -1
                        }
                    }
                    $after = [string]$range.Text
                    if ($after.TrimEnd([char[]]@([char]13, [char]10)) -cne $values[$cellIndex]) {
                        throw ('Native table cell did not retain its authored text: ' + $profile.Name + '/' + $cellIndex)
                    }
                    $observed.Add([pscustomobject]@{ cell = $cellIndex; beforeText = $before; afterText = $after; beforeLanguageId = $beforeLanguage; afterLanguageId = [int]$range.LanguageID })
                }
            }
        }
        $expected = [pscustomobject]@{ name = $profile.Name; kind = 'Table'; rows = 2; columns = 2; cells = $values; left = $left; top = $top; width = $width; height = $height }
        if ($Case -ne 'TableLabel') {
            $expected | Add-Member -NotePropertyName cellTyping -NotePropertyValue ([pscustomobject]@{
                profile = $profile; observed = $observed.ToArray(); rangeCapture = $(if ($Case -eq 'TableCellTypingEarlyRange') { 'BeforeText' } else { 'FreshWholeCell' })
            })
        }
        $expectedShapes.Add($expected)
        foreach ($value in $values) { if ($value.Length -gt 0) { $expectedText.Add($value) } }
    }
    if ($Case -eq 'TableLabel') {
        New-ReferencePicture $PicturePath
        $image = $FirstPage.Shapes.AddPicture($PicturePath, 0, -1, 216, 72, 54, 54)
        $image.Name = 'LabelImage'; $objects.Add($image)
        $expectedShapes.Add([pscustomobject]@{ name = 'LabelImage'; kind = 'Image'; left = 216; top = 72; width = 54; height = 54 })
    }
    return $pages
}
