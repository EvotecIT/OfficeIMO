# Retained text identifies formatting edits independently of destination styles.
function Get-NativeFormattingTopologyProfiles {
    $base = @('First beta tail', 'Second gamma end', 'Third delta finish', 'Final epsilon close')
    $profiles = @(
        @{ Name = 'Baseline'; Order = @(0, 1, 2, 3); FrenchRunParagraph = 1 },
        @{ Name = 'DropEmphasis'; Order = @(0, 2, 3); RunOverrides = @{ 0 = @{ 0 = @{ Bold = 0 }; 1 = @{ Italic = 0 } } } },
        @{ Name = 'ReorderSizes'; Order = @(3, 1, 0); FrenchRunParagraph = 1; RunOverrides = @{ 0 = @{ 0 = @{ Size = 16 }; 1 = @{ Size = 10 }; 2 = @{ Size = 18 } } } },
        @{ Name = 'DuplicateColors'; Order = @(0, 1, 1, 2, 3); FrenchRunParagraph = @(1, 2); RunOverrides = @{ 1 = @{ 0 = @{ Rgb = 32768 }; 1 = @{ Rgb = 16711680 } }; 2 = @{ 0 = @{ Family = 'Cambria' } } } },
        @{ Name = 'OnlyFinal'; Order = @(3); RunOverrides = @{ 0 = @{ 0 = @{ Bold = 0 }; 1 = @{ Italic = 0 }; 2 = @{ Size = 18; Bold = -1; Rgb = 32768 } } } },
        @{ Name = 'OnlyFirstLayout'; Order = @(0); ParagraphOverrides = @{ 0 = @{ Alignment = 1; SpaceBefore = 4; SpaceAfter = 9; LeftIndent = 12; FirstLineIndent = 6 } } },
        @{ Name = 'Combined'; Order = @(2, 0, 3); ParagraphOverrides = @{ 0 = @{ Alignment = 0; SpaceAfter = 11 } }; RunOverrides = @{ 0 = @{ 0 = @{ Family = 'Cambria'; Size = 16; Rgb = 32768; Bold = 0; Underline = 3 }; 1 = @{ Size = 10 } } } },
        @{ Name = 'DestinationFormatCollision'; Order = @(0, 3); ParagraphOverrides = @{ 0 = @{ Alignment = 1; SpaceBefore = 2; SpaceAfter = 6 } } },
        @{ Name = 'AmbiguousTextLanguage'; Order = @(0, 0, 3); FrenchRunParagraph = 1 },
        @{ Name = 'RetainedEnglishWithFinalFormat'; Order = @(0, 3); RunOverrides = @{ 1 = @{ 0 = @{ Size = 16 } } } },
        @{ Name = 'RetainedFrenchWithFinalFormat'; Order = @(0, 3); FrenchRunParagraph = 0; RunOverrides = @{ 1 = @{ 0 = @{ Size = 16 } } } },
        @{ Name = 'EquivalentSource'; Order = @(0, 0, 3) },
        @{ Name = 'EquivalentFormatting'; Order = @(0, 3); RunOverrides = @{ 0 = @{ 1 = @{ Italic = 0 } } } },
        @{ Name = 'FirstRunOnly'; Order = @(0, 2, 3); RunOverrides = @{ 0 = @{ 0 = @{ Size = 18 } } } },
        @{ Name = 'ParagraphOnly'; Order = @(0, 2, 3); ParagraphOverrides = @{ 0 = @{ SpaceAfter = 11 } } },
        @{ Name = 'ResourceGrowth'; Order = @(2, 0, 3); ParagraphOverrides = @{ 0 = @{ Alignment = 0; SpaceAfter = 11 } }; RunOverrides = @{ 0 = @{ 0 = @{ Family = 'Consolas'; Size = 16; Rgb = 32768; Bold = 0; Underline = 3 }; 1 = @{ Size = 10 } } } }
    )
    foreach ($profile in $profiles) {
        $profile.MixedRuns = $true
        $profile.LiteralText = ($profile.Order | ForEach-Object { $base[$_] }) -join "`r"
        $profile
    }
}
