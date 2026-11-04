function Add-NativeComparisonLayout {
    <#
    .SYNOPSIS
    Authors the brochure and master-page comparison layouts in a new publication.
    .DESCRIPTION
    Called only by the isolated native worker after process ownership is recorded.
    Uses its explicit Arial text-frame helper and authored-content collections.
    Returns the declared layout and master assignments, which the worker verifies
    against native observations before saving. No existing publication is opened.
    .PARAMETER Case
    Brochure, MultipleMasters, TwoPageMaster or the edited master variants.
    .PARAMETER Document
    The task-owned, newly created Microsoft Publisher publication.
    .PARAMETER FirstPage
    Its first physical page.
    .OUTPUTS
    The authored page count, panel width and native master assignment IDs.
    #>
    param(
        [ValidateSet('Brochure', 'MultipleMasters', 'TwoPageMaster', 'MultipleMastersEdited', 'TwoPageMasterEdited', 'MultipleMastersReassigned', 'MultipleMastersIgnored', 'TwoPageMasterReassigned', 'TwoPageMasterIgnored')]
        [string] $Case,
        $Document,
        $FirstPage
    )

    $edited = $Case.EndsWith('Edited', [StringComparison]::Ordinal)
    $baseCase = $Case -replace '(Edited|Reassigned|Ignored)$', ''
    $reassigned = $Case.EndsWith('Reassigned', [StringComparison]::Ordinal)
    $ignored = $Case.EndsWith('Ignored', [StringComparison]::Ordinal)
    $moveX = if ($edited) { 6 } else { 0 }
    $moveY = if ($edited) { 12 } else { 0 }
    $suffix = if ($edited) { ' updated' } else { '' }
    if ($Case -eq 'Brochure') {
        Set-NativePhase 'Brochure-AddSecondPage'
        $second = $Document.Pages.Add(1, 1)
        $objects.Add($second)
        $sides = @($FirstPage, $second)
        for ($side = 0; $side -lt 2; $side++) {
            for ($panel = 0; $panel -lt 3; $panel++) {
                $left = 18 + $panel * 264
                $name = 'BrochureSide' + ($side + 1) + 'Panel' + ($panel + 1)
                $title = 'Side ' + ($side + 1) + ', panel ' + ($panel + 1)
                $body = 'This independent brochure panel uses explicit Arial text and a fixed 228-point column. Its placement distinguishes the outside and inside of a three-panel publication.'
                Set-NativePhase ($name + '-AddTitle')
                $null = Add-TextFrame $sides[$side] ($name + 'Title') $title $left 36 228 36 18
                Set-NativePhase ($name + '-AddBody')
                $null = Add-TextFrame $sides[$side] ($name + 'Body') ($title + '. ' + $body) $left 96 228 420 12
                Set-NativePhase ($name + '-AddMarker')
                $marker = $sides[$side].Shapes.AddShape(1, $left, 78, 228, 6)
                $marker.Name = $name + 'Marker'
                $marker.Line.Visible = 0
                $marker.Fill.ForeColor.RGB = 10506797
                $objects.Add($marker)
                $expectedShapes.Add([pscustomobject]@{ name = $marker.Name; kind = 'Rectangle'; pageIndex = $side + 1; left = $left; top = 78; width = 228; height = 6 })
            }
        }
        return [pscustomobject]@{ pages = 2; panelWidthPoints = 264; masterAssignments = @() }
    }

    Set-NativePhase ($Case + '-AddSecondPage')
    $second = $Document.Pages.Add(1, 1)
    Set-NativePhase ($Case + '-AddThirdPage')
    $third = $Document.Pages.Add(1, 2)
    $objects.Add($second)
    $objects.Add($third)
    $physicalPages = @($FirstPage, $second, $third)
    Set-NativePhase ($Case + '-ObserveDefaultMaster')
    $defaultMaster = $Document.MasterPages.Item(1)
    $objects.Add($defaultMaster)
    if ($Document.MasterPages.Count -ne 1 -or $defaultMaster.IsTwoPageMaster) {
        throw 'The master control requires the new publication to have one single-page default master.'
    }

    if ($baseCase -eq 'MultipleMasters') {
        Set-NativePhase 'MultipleMasters-AddAlternateMaster'
        $alternateMaster = $Document.MasterPages.Add($false, 'B', 'Comparison Master B')
        $objects.Add($alternateMaster)
        Set-NativePhase 'MultipleMasters-AddAlphaMarker'
        $null = Add-TextFrame $defaultMaster 'MasterAlphaMarker' ('Master Alpha' + $suffix) (18 + $moveX) (24 + $moveY) 252 24 10
        Set-NativePhase 'MultipleMasters-AddBetaMarker'
        $null = Add-TextFrame $alternateMaster 'MasterBetaMarker' ('Master Beta' + $suffix) (342 + $moveX) (24 + $moveY) 252 24 10
        $assignedMasters = if ($reassigned) { @($alternateMaster, $defaultMaster, $alternateMaster) }
            else { @($defaultMaster, $alternateMaster, $defaultMaster) }
    } else {
        $beforeIds = @([int]$defaultMaster.PageID)
        # Add returns one member without documenting which side. Resolve both
        # members by their newly observed IDs and native collection order.
        # The qualified Publisher 16.0 build accepts P but rejects TP with
        # E_INVALIDARG. Its API documentation does not state a maximum length.
        Set-NativePhase 'TwoPageMaster-AddPair'
        $createdPair = $Document.MasterPages.Add($true, 'P', 'Comparison Two-Page Master')
        $objects.Add($createdPair)
        $newMasters = [Collections.Generic.List[object]]::new()
        for ($index = 1; $index -le $Document.MasterPages.Count; $index++) {
            Set-NativePhase ('TwoPageMaster-ObservePairEntry-' + $index)
            $master = $Document.MasterPages.Item($index)
            $objects.Add($master)
            if ([int]$master.PageID -notin $beforeIds) {
                $newMasters.Add($master)
            }
        }
        Set-NativePhase 'TwoPageMaster-CheckPairFlags'
        if ($newMasters.Count -ne 2 -or -not $newMasters[0].IsTwoPageMaster -or -not $newMasters[1].IsTwoPageMaster) {
            throw 'Publisher did not create two distinct members of a native two-page master.'
        }
        Set-NativePhase 'TwoPageMaster-AddDefaultMarker'
        $null = Add-TextFrame $defaultMaster 'DefaultMasterMarker' ('Default master' + $suffix) (18 + $moveX) (24 + $moveY) 252 24 10
        Set-NativePhase 'TwoPageMaster-AddLeftMarker'
        $null = Add-TextFrame $newMasters[0] 'LeftMasterMarker' ('Paired master left' + $suffix) (18 + $moveX) (24 + $moveY) 252 24 10
        Set-NativePhase 'TwoPageMaster-AddRightMarker'
        $null = Add-TextFrame $newMasters[1] 'RightMasterMarker' ('Paired master right' + $suffix) (342 + $moveX) (24 + $moveY) 252 24 10
        $assignedMasters = if ($reassigned) { @($newMasters[1], $newMasters[0], $defaultMaster) }
            else { @($defaultMaster, $newMasters[0], $newMasters[1]) }
    }

    $assignments = [Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt 3; $index++) {
        Set-NativePhase ($Case + '-AssignPhysicalMaster-' + ($index + 1))
        $physicalPages[$index].Master = $assignedMasters[$index]
        if ($ignored -or $reassigned) {
            Set-NativePhase ($Case + '-SetIgnoreMaster-' + ($index + 1))
            $suppress = $ignored -and $index -ne 1
            $physicalPages[$index].IgnoreMaster = $suppress
            if ([bool]$physicalPages[$index].IgnoreMaster -ne $suppress) {
                throw 'Publisher did not retain the requested IgnoreMaster state.'
            }
        }
        Set-NativePhase ($Case + '-AddPhysicalText-' + ($index + 1))
        $null = Add-TextFrame $physicalPages[$index] ('PageIdentity' + ($index + 1)) ($baseCase + ' page ' + ($index + 1)) (72 + $moveX) (96 + $moveY) 468 48 16
        $assignments.Add([pscustomobject]@{
            pageIndex = $index + 1
            masterPageId = [int]$assignedMasters[$index].PageID
        })
    }
    return [pscustomobject]@{ pages = 3; panelWidthPoints = $null; masterAssignments = $assignments.ToArray() }
}
