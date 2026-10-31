# Each profile is authored independently, rather than mutating the baseline.
function Get-NativeParagraphTopologyProfiles {
    param([switch] $EmptyParagraphs, [switch] $SizeWitnesses, [switch] $UniformEmpty, [switch] $TextReplacements)
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
    if ($TextReplacements) {
        $profiles = @(
            @{ Name = 'Baseline'; Order = @(0, 1, 2, 3) },
            @{ Name = 'GrowDrop'; Order = @(0, 2, 3); LiteralText = "First native paragraph expanded`rThird updated`rFinal native paragraph" },
            @{ Name = 'ShrinkDrop'; Order = @(1, 2, 3); LiteralText = "Second short`rThird short`rFinal native paragraph" },
            @{ Name = 'RewriteReorder'; Order = @(3, 1, 0); LiteralText = "Final edited paragraph`rSecond updated`rFirst updated" },
            @{ Name = 'RewriteDuplicate'; Order = @(0, 1, 1, 2, 3); LiteralText = "First native paragraph`rFirst centered edit`rSecond centered edit`rThird native paragraph`rFinal native paragraph" },
            @{ Name = 'OnlyFinal'; Order = @(3); LiteralText = 'Final edited paragraph' },
            @{ Name = 'OnlyFirst'; Order = @(0); LiteralText = 'First Ω edited' },
            @{ Name = 'AmbiguousLanguage'; Order = @(0, 0, 3); FrenchParagraph = 1 },
            @{ Name = 'EquivalentSource'; Order = @(0, 0, 3); LiteralText = "First shared-style source`rSecond shared-style source`rFinal native paragraph" },
            @{ Name = 'EquivalentReplacement'; Order = @(0, 3); LiteralText = "Shared-style replacement`rFinal native paragraph" },
            @{ Name = 'EmptyBaseline'; Order = @(0, 1, 2, 3); LiteralText = "First native paragraph`r`rThird native paragraph`r" },
            @{ Name = 'RewriteWithRetainedEmpty'; Order = @(2, 3); LiteralText = "Third updated`r" },
            @{ Name = 'RetainedEnglishWithRewrite'; Order = @(0, 3); LiteralText = "First native paragraph`rFinal edited paragraph" },
            @{ Name = 'RetainedFrenchWithRewrite'; Order = @(0, 3); LiteralText = "First native paragraph`rFinal edited paragraph"; FrenchParagraph = 0 }
        )
    } elseif ($UniformEmpty) {
        $texts[1] = ''
        $profiles = @(
            @{ Name = 'UniformEmptyLanguage'; Order = @(1, 1); FrenchParagraph = 1 },
            @{ Name = 'UniformRetainedEnglish'; Order = @(1) },
            @{ Name = 'UniformRetainedFrench'; Order = @(1); FrenchParagraph = 0 },
            @{ Name = 'UniformMixedEnglishFinal'; Order = @(1, 1, 1); LiteralText = "First authored paragraph`rSecond authored paragraph`r" },
            @{ Name = 'UniformMixedFrenchInterior'; Order = @(1, 1, 1); LiteralText = "`rFirst authored paragraph`rSecond authored paragraph"; FrenchParagraph = 0 }
        )
    } elseif ($EmptyParagraphs) {
        $texts[1] = ''
        $texts[3] = ''
        $profiles += @(
            @{ Name = 'InteriorEmptyLast'; Order = @(0, 1) },
            @{ Name = 'FinalEmptyOnly'; Order = @(3) },
            @{ Name = 'TwoEmpty'; Order = @(1, 3) },
            @{ Name = 'AmbiguousEmptyLanguage'; Order = @(1, 1, 3); FrenchParagraph = 1 }
        )
        if ($SizeWitnesses) {
            $profiles += @(
                @{ Name = 'StyleSize12Witness'; Order = @(1); LiteralText = 'Native style size witness'; WitnessStyleSize = 12 },
                @{ Name = 'StyleSize10Witness'; Order = @(1); LiteralText = 'Native style size witness'; WitnessStyleSize = 10 },
                @{ Name = 'StyleSize14Witness'; Order = @(1); LiteralText = 'Native style size witness'; WitnessStyleSize = 14 },
                @{ Name = 'StyleSize12Bi9Witness'; Order = @(1); LiteralText = 'Native style size witness'; WitnessStyleKey = 'size12bi9' },
                @{ Name = 'DirectSize14Bi9Witness'; Order = @(0); LiteralText = 'Native style size witness'; DirectSize = 14; DirectRtlSize = 9 }
            )
        }
    }
    foreach ($profile in $profiles) {
        $profile.Text = ($profile.Order | ForEach-Object { $texts[$_] }) -join "`r"
        if ($profile.ContainsKey('LiteralText')) { $profile.Text = $profile.LiteralText }
        $profile
    }
}
