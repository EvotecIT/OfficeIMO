function Set-NativeCandidateText {
    <#
    .SYNOPSIS
    Changes one qualified text range in the isolated writable candidate snapshot.
    .PARAMETER Publication
    The task-owned snapshot opened by the native validation worker.
    .PARAMETER Objects
    Retains COM wrappers for the worker's lifetime.
    .PARAMETER Target
    The sole unlinked text box, or the first unmerged cell of the sole physical table.
    .PARAMETER Text
    The replacement text supplied to the controller.
    .OUTPUTS
    Target identity and before/requested/observed text for the validation report.
    #>
    param($Publication, [Collections.Generic.List[object]] $Objects,
        [ValidateSet('TextBox', 'FirstTableCell')] [string] $Target,
        [AllowEmptyString()] [string] $Text)

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
        if ($cell.Width -ne 1 -or $cell.Height -ne 1) { throw 'Native cell mutation requires an unmerged first cell.' }
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
