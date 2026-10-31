# Independent final-state controls use paragraph formatting without selection or clipboard access.
function Add-NativeParagraphTopologyControl {
    param($Publication, $FirstPage, [switch] $EmptyParagraphs, [switch] $UniformEmpty)
    $profiles = @(Get-NativeParagraphTopologyProfiles -EmptyParagraphs:$EmptyParagraphs -SizeWitnesses:$EmptyParagraphs -UniformEmpty:$UniformEmpty)
    if ($EmptyParagraphs) {
        $typingStyles = Add-NativeEmptyParagraphTypingStyles -Publication $Publication
        foreach ($profile in $profiles) { $profile.EmptyTypingStyles = $typingStyles }
    }
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
        if ($profile.ContainsKey('DirectSize')) {
            $whole.Font.Size = $profile.DirectSize
            $whole.Font.SizeBi = $profile.DirectRtlSize
            $whole.Font.Bold = 0
            $whole.Font.Italic = 0
            $whole.Font.BoldBi = -1
            $whole.Font.ItalicBi = -1
            $expectedShapes[$expectedShapes.Count - 1].size = $profile.DirectSize
        }
        $observed = @(Set-NativeParagraphTopologyFormatting -WholeRange $whole -Profile $profile)
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName paragraphTopology -NotePropertyValue ([pscustomobject]@{
            order = $profile.Order
            observed = $observed
            typingStyles = $(if ($EmptyParagraphs) { $typingStyles.observations } else { $null })
        })
    }
    return $profiles.Count
}
