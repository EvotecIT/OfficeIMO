# Three-word paragraphs expose independent run-length and native-template changes.
function Get-NativeMixedRunTextProfiles {
    $profiles = @(
        @{ Name = 'Baseline'; Order = @(0, 1, 2, 3); LiteralText = "First beta tail`rSecond gamma end`rThird delta finish`rFinal epsilon close" },
        @{ Name = 'GrowDrop'; Order = @(0, 2, 3); LiteralText = "Expanded lengthened ending`rUpdated rewritten result`rFinal epsilon close" },
        @{ Name = 'ShrinkDrop'; Order = @(1, 2, 3); LiteralText = "S b t`rT c z`rFinal epsilon close" },
        @{ Name = 'RewriteReorder'; Order = @(3, 1, 0); LiteralText = "Finale changed ending`rSecond altered tail`rFirst edited finish" },
        @{ Name = 'RewriteDuplicate'; Order = @(0, 1, 1, 2, 3); LiteralText = "First beta tail`rFirst centered edit`rSecond centered edit`rThird delta finish`rFinal epsilon close" },
        @{ Name = 'OnlyFinal'; Order = @(3); LiteralText = 'Finale modified ending' },
        @{ Name = 'OnlyFirst'; Order = @(0); LiteralText = 'Ωmega replaced finish' },
        @{ Name = 'AmbiguousLanguage'; Order = @(0, 0, 3); LiteralText = "First beta tail`rFirst beta tail`rFinal epsilon close"; FrenchRunParagraph = 1 },
        @{ Name = 'EquivalentSource'; Order = @(0, 0, 3); LiteralText = "First beta tail`rAnother gamma end`rFinal epsilon close" },
        @{ Name = 'EquivalentReplacement'; Order = @(0, 3); LiteralText = "Common edited result`rFinal epsilon close" },
        @{ Name = 'RetainedEnglishWithRewrite'; Order = @(0, 3); LiteralText = "First beta tail`rFinale modified ending" },
        @{ Name = 'RetainedFrenchWithRewrite'; Order = @(0, 3); LiteralText = "First beta tail`rFinale modified ending"; FrenchRunParagraph = 0 },
        @{ Name = 'MiddleRunOnly'; Order = @(0, 2, 3); LiteralText = "First altered tail`rThird delta finish`rFinal epsilon close" }
    )
    foreach ($profile in $profiles) { $profile.MixedRuns = $true; $profile }
}
