# Uses documented paragraph ranges; returned observations include native defaults.
function Set-NativeParagraphTopologyFormatting {
    param($WholeRange, $Profile)
    for ($paragraph = 0; $paragraph -lt $Profile.Order.Count; $paragraph++) {
        $range = $WholeRange.Paragraphs($paragraph + 1, 1)
        $objects.Add($range)
        $format = $range.ParagraphFormat
        $objects.Add($format)
        switch ($Profile.Order[$paragraph]) {
            0 { $format.Alignment = 0; $format.SpaceAfter = 3 }
            1 { $format.Alignment = 1; $format.SpaceBefore = 2; $format.SpaceAfter = 6 }
            2 { $format.Alignment = 2; $format.SpaceBefore = 5; $format.SpaceAfter = 2 }
            3 { $format.Alignment = 0; $format.LeftIndent = 18; $format.FirstLineIndent = 6; $format.SpaceBefore = 7; $format.SpaceAfter = 4 }
        }
        if ($Profile.ContainsKey('FrenchParagraph') -and $paragraph -eq $Profile.FrenchParagraph) { $range.LanguageID = 1036 }
        [pscustomobject]@{
            sourceOrdinal = $Profile.Order[$paragraph]
            text = [string]$range.Text
            alignment = [int]$format.Alignment
            spaceBefore = [double]$format.SpaceBefore
            spaceAfter = [double]$format.SpaceAfter
            leftIndent = [double]$format.LeftIndent
            firstLineIndent = [double]$format.FirstLineIndent
            languageId = [int]$range.LanguageID
        }
    }
}
