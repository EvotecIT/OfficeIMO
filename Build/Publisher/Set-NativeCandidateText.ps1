function Set-NativeCandidateText {
    <#
    .SYNOPSIS
    Changes one qualified text range in the isolated writable candidate snapshot.
    .PARAMETER Publication
    The task-owned snapshot opened by the native validation worker.
    .PARAMETER Objects
    Retains COM wrappers for the worker's lifetime.
    .PARAMETER Target
    The sole unlinked text box, the first unmerged cell of the sole physical table,
    its first horizontal merged anchor with exactly two columns and one row,
    or one first character of the qualified three-frame page 1/3/2 story.
    .PARAMETER Text
    The replacement text supplied to the controller. LinkedStoryFirstCharacter
    accepts one printable ASCII character distinct from the original M.
    .OUTPUTS
    Target identity and before/requested/observed text for the validation report.
    #>
    param($Publication, [Collections.Generic.List[object]] $Objects,
        [ValidateSet('TextBox', 'FirstTableCell', 'FirstMergedTableCell', 'LinkedStoryFirstCharacter')] [string] $Target,
        [AllowEmptyString()] [string] $Text)

    if ($Target -eq 'LinkedStoryFirstCharacter') {
        if ($Publication.ReadOnly -or $Text.Length -ne 1 -or $Text[0] -lt [char]0x20 -or
            $Text[0] -gt [char]0x7e -or $Text -ceq 'M') {
            throw 'Linked story mutation requires a writable snapshot and one printable ASCII replacement character distinct from M.'
        }
        $beforeState = Get-NativeLinkedStoryEditState -Publication $Publication -Objects $Objects
        $before = $beforeState.Snapshot.text
        $expected = $Text + $before.Substring(1)
        $character = $beforeState.Range.Characters(1, 1)
        $Objects.Add($character)
        if ([string]$character.Text -cne 'M' -or $character.Length -ne 1) {
            throw 'The linked story mutation did not select only its original first M character.'
        }
        $character.Text = $Text
        $afterState = Get-NativeLinkedStoryEditState -Publication $Publication -Objects $Objects -ExpectedText $expected
        Assert-NativeLinkedStoryEditState -Before $beforeState.Snapshot -After $afterState.Snapshot
        return [ordered]@{
            target = $Target
            shapeName = $beforeState.ShapeName
            before = $before
            requested = $Text
            observed = $afterState.Snapshot.text
            beforeState = $beforeState.Snapshot
            afterState = $afterState.Snapshot
            reopenedState = $null
        }
    }
    if ($Publication.ReadOnly -or $Publication.Pages.Count -ne 1) {
        throw 'Native text mutation requires one writable physical page.'
    }
    $page = $Publication.Pages.Item(1)
    $Objects.Add($page)
    if ($Target -eq 'TextBox') {
        if ($page.Shapes.Count -ne 1) { throw 'Native text mutation requires exactly one physical text box.' }
        $shape = $page.Shapes.Item(1)
        $Objects.Add($shape)
        if ($shape.Type -ne 17 -or $shape.HasTextFrame -ne -1) { throw 'Native text mutation requires an ordinary text box.' }
        $frame = $shape.TextFrame
        $Objects.Add($frame)
        if ($frame.HasNextLink -eq -1 -or $frame.HasPreviousLink -eq -1) { throw 'Native text mutation does not accept linked frames.' }
        $range = $frame.TextRange
    } else {
        $tables = [Collections.Generic.List[object]]::new()
        for ($index = 1; $index -le $page.Shapes.Count; $index++) {
            $candidate = $page.Shapes.Item($index)
            $Objects.Add($candidate)
            if ($candidate.Type -eq 18) { $tables.Add($candidate) }
        }
        if ($tables.Count -ne 1) { throw 'Native cell mutation requires exactly one physical table.' }
        $shape = $tables[0]
        $table = $shape.Table
        $cells = $table.Cells(1, 1, 1, 1)
        $cell = $cells.Item(1)
        $Objects.Add($table)
        $Objects.Add($cells)
        $Objects.Add($cell)
        if ($Target -eq 'FirstTableCell') {
            if ($cell.Width -ne 1 -or $cell.Height -ne 1) { throw 'Native cell mutation requires an unmerged first cell.' }
        } elseif ($Target -eq 'FirstMergedTableCell') {
            if ($cell.Width -ne 2 -or $cell.Height -ne 1) { throw 'Native merged-cell mutation requires a two-column, one-row first anchor.' }
        }
        $range = $cell.TextRange
    }
    $Objects.Add($range)
    $before = [string]$range.Text
    $range.Text = $Text
    $observed = [string]$range.Text
    # Publisher can materialize the final paragraph CR only when saving.
    if ($observed -cne $Text -and $observed -cne ($Text + "`r")) {
        throw 'Publisher did not retain the requested native text mutation.'
    }
    return [ordered]@{ target = $Target; shapeName = [string]$shape.Name; before = $before; requested = $Text; observed = $observed }
}
