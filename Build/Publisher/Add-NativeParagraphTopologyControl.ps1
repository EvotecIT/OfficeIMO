# Independent final-state controls use paragraph formatting, without selection,
# clipboard access or character/font binding probes.
function Add-NativeParagraphTopologyControl {
    param($Publication, $FirstPage)
    $texts = @('First native paragraph', 'Second native paragraph', 'Third native paragraph', 'Final native paragraph')
    $profiles = @(
        @{ Name = 'Baseline'; Order = @(0, 1, 2, 3) },
        @{ Name = 'DropFirst'; Order = @(1, 2, 3) },
        @{ Name = 'DropMiddle'; Order = @(0, 2, 3) },
        @{ Name = 'DropFinal'; Order = @(0, 1, 2) },
        @{ Name = 'Reorder'; Order = @(3, 0, 2) },
        @{ Name = 'Duplicate'; Order = @(0, 1, 1, 2, 3) },
        @{ Name = 'AmbiguousLanguage'; Order = @(0, 0, 3); FrenchParagraph = 1 }
    )
    $owner = $FirstPage
    for ($index = 0; $index -lt $profiles.Count; $index++) {
        $profile = $profiles[$index]
        if ($index -gt 0) {
            $owner = $Publication.Pages.Add(1, $index)
            $objects.Add($owner)
        }
        Set-NativePhase ('ParagraphTopology-' + $profile.Name)
        $text = ($profile.Order | ForEach-Object { $texts[$_] }) -join "`r"
        $shape = Add-TextFrame $owner $profile.Name $text 72 72 468 216 12
        $whole = $shape.TextFrame.TextRange
        $objects.Add($whole)
        $observed = [Collections.Generic.List[object]]::new()
        for ($paragraph = 0; $paragraph -lt $profile.Order.Count; $paragraph++) {
            $range = $whole.Paragraphs($paragraph + 1, 1)
            $objects.Add($range)
            $format = $range.ParagraphFormat
            $objects.Add($format)
            switch ($profile.Order[$paragraph]) {
                0 { $format.Alignment = 0; $format.SpaceAfter = 3 }
                1 { $format.Alignment = 1; $format.SpaceBefore = 2; $format.SpaceAfter = 6 }
                2 { $format.Alignment = 2; $format.SpaceBefore = 5; $format.SpaceAfter = 2 }
                3 { $format.Alignment = 0; $format.LeftIndent = 18; $format.FirstLineIndent = 6; $format.SpaceBefore = 7; $format.SpaceAfter = 4 }
            }
            if ($profile.ContainsKey('FrenchParagraph') -and $paragraph -eq $profile.FrenchParagraph) { $range.LanguageID = 1036 }
            $observed.Add([pscustomobject]@{
                sourceOrdinal = $profile.Order[$paragraph]
                text = [string]$range.Text
                alignment = [int]$format.Alignment
                spaceBefore = [double]$format.SpaceBefore
                spaceAfter = [double]$format.SpaceAfter
                leftIndent = [double]$format.LeftIndent
                firstLineIndent = [double]$format.FirstLineIndent
                languageId = [int]$range.LanguageID
            })
        }
        $expectedShapes[$expectedShapes.Count - 1] | Add-Member -NotePropertyName paragraphTopology -NotePropertyValue ([pscustomobject]@{
            order = $profile.Order
            observed = $observed.ToArray()
        })
    }
    return $profiles.Count
}
