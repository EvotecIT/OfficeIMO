# Uses documented paragraph ranges; returned observations include native defaults.
function Set-NativeParagraphTopologyFormatting {
    param($WholeRange, $Profile)
    for ($paragraph = 0; $paragraph -lt $Profile.Order.Count; $paragraph++) {
        $range = $WholeRange.Paragraphs($paragraph + 1, 1)
        $objects.Add($range)
        $format = $range.ParagraphFormat
        $objects.Add($format)
        if ($Profile.ContainsKey('EmptyTypingStyles')) {
            $key = if ($Profile.ContainsKey('WitnessStyleKey')) { $Profile.WitnessStyleKey }
                elseif ($Profile.ContainsKey('WitnessStyleSize')) { 'size' + $Profile.WitnessStyleSize } else { [string]$Profile.Order[$paragraph] }
            if ($Profile.EmptyTypingStyles.ContainsKey($key)) { $format.TextStyle = $Profile.EmptyTypingStyles[$key] }
        }
        switch ($Profile.Order[$paragraph]) {
            0 { $format.Alignment = 0; $format.SpaceAfter = 3 }
            1 { $format.Alignment = 1; $format.SpaceBefore = 2; $format.SpaceAfter = 6 }
            2 { $format.Alignment = 2; $format.SpaceBefore = 5; $format.SpaceAfter = 2 }
            3 { $format.Alignment = 0; $format.LeftIndent = 18; $format.FirstLineIndent = 6; $format.SpaceBefore = 7; $format.SpaceAfter = 4 }
        }
        if ($Profile.ContainsKey('ParagraphOverrides') -and $Profile.ParagraphOverrides.ContainsKey($paragraph)) {
            $override = $Profile.ParagraphOverrides[$paragraph]
            foreach ($name in 'Alignment', 'SpaceBefore', 'SpaceAfter', 'LeftIndent', 'FirstLineIndent') {
                if ($override.ContainsKey($name)) { $format.$name = $override[$name] }
            }
        }
        if ($Profile.ContainsKey('FrenchParagraph') -and $paragraph -eq $Profile.FrenchParagraph) { $range.LanguageID = 1036 }
        $runFormatting = if ($Profile.ContainsKey('MixedRuns')) {
            $overrides = if ($Profile.ContainsKey('RunOverrides') -and $Profile.RunOverrides.ContainsKey($paragraph)) { $Profile.RunOverrides[$paragraph] } else { @{} }
            @(Set-NativeMixedRunFormatting -ParagraphRange $range -FrenchMiddle:($Profile.ContainsKey('FrenchRunParagraph') -and @($Profile.FrenchRunParagraph) -contains $paragraph) -Overrides $overrides)
        } else { $null }
        [pscustomobject]@{
            sourceOrdinal = $Profile.Order[$paragraph]
            text = [string]$range.Text
            alignment = [int]$format.Alignment
            spaceBefore = [double]$format.SpaceBefore
            spaceAfter = [double]$format.SpaceAfter
            leftIndent = [double]$format.LeftIndent
            firstLineIndent = [double]$format.FirstLineIndent
            languageId = [int]$range.LanguageID
            textStyle = [string]$format.TextStyle
            mixedRuns = $runFormatting
        }
    }
}
