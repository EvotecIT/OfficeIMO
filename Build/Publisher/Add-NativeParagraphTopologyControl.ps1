# Independent final-state controls use paragraph formatting without selection or clipboard access.
function Add-NativeParagraphTopologyControl {
    param($Publication, $FirstPage)
    $profiles = @(Get-NativeParagraphTopologyProfiles)
    $owner = $FirstPage
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0) {
            $owner = $Publication.Pages.Add(1, $index)
            $objects.Add($owner)
        }
        Set-NativePhase ('ParagraphTopology-' + $profile.Name)
        $shape = Add-TextFrame $owner $profile.Name $profile.Text 72 72 468 216 12
        $whole = $shape.TextFrame.TextRange
        $objects.Add($whole)
        $observed = @(Set-NativeParagraphTopologyFormatting -WholeRange $whole -Profile $profile)
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName paragraphTopology -NotePropertyValue ([pscustomobject]@{
            order = $profile.Order
            observed = $observed
        })
    }
    return $profiles.Count
}
