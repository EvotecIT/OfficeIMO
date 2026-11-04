# Native word ranges expose visible run changes while retaining hidden language boundaries.
function Get-NativeRunSegmentationProfiles {
    $texts = @('First beta tail', 'Second gamma end', 'Third delta finish', 'Final epsilon close')
    $normal = @{}
    $mono = @{}
    foreach ($word in 0, 1, 2) {
        $normal[$word] = @{ Size = 12; Bold = 0; Italic = 0; Rgb = 0 }
        $mono[$word] = @{ Size = 16; Bold = 0; Italic = 0; Family = 'Consolas'; Rgb = 32768 }
    }
    $two = @{ 0 = @{ Size = 14; Bold = -1; Italic = 0; Rgb = 16711680 }; 1 = @{ Size = 14; Bold = -1; Italic = 0; Rgb = 16711680 } }
    $profiles = @(
        @{ Name = 'MixedBaseline'; Order = @(0, 1, 2, 3); FrenchRunParagraph = 1 },
        @{ Name = 'MergeDrop'; Order = @(0, 2, 3); RunOverrides = @{ 0 = $normal } },
        @{ Name = 'MergeReorder'; Order = @(3, 1, 0); FrenchRunParagraph = 1; RunOverrides = @{ 0 = $mono } },
        @{ Name = 'MergeDuplicate'; Order = @(0, 1, 1, 2, 3); FrenchRunParagraph = @(1, 2); RunOverrides = @{ 1 = $normal; 2 = $mono } },
        @{ Name = 'MergeOnlyFinal'; Order = @(3); RunOverrides = @{ 0 = $normal } },
        @{ Name = 'MergeFirstTwo'; Order = @(0, 2, 3); RunOverrides = @{ 0 = $two } },
        @{ Name = 'UniformBaseline'; Order = @(0, 1, 2, 3); Uniform = $true; FrenchRunParagraph = 1 },
        @{ Name = 'SplitDrop'; Order = @(0, 2, 3); Uniform = $true; MixedParagraphs = @(0) },
        @{ Name = 'SplitReorder'; Order = @(3, 1, 0); Uniform = $true; MixedParagraphs = @(0); FrenchRunParagraph = 1 },
        @{ Name = 'SplitDuplicate'; Order = @(0, 1, 1, 2, 3); Uniform = $true; MixedParagraphs = @(1); FrenchRunParagraph = @(1, 2); RunOverrides = @{ 2 = $two } },
        @{ Name = 'SplitOnlyFinal'; Order = @(3); Uniform = $true; MixedParagraphs = @(0) },
        @{ Name = 'AmbiguousMixedLanguage'; Order = @(0, 0, 3); FrenchRunParagraph = 1 },
        @{ Name = 'EquivalentMixedTemplates'; Order = @(0, 0, 3) },
        @{ Name = 'MergeEquivalent'; Order = @(0, 3); RunOverrides = @{ 0 = $normal } },
        @{ Name = 'AmbiguousUniformLanguage'; Order = @(0, 0, 3); Uniform = $true; FrenchRunParagraph = 1 },
        @{ Name = 'MergeHiddenLanguage'; Order = @(1, 3); FrenchRunParagraph = 0; RunOverrides = @{ 0 = $normal } },
        @{ Name = 'SplitHiddenLanguage'; Order = @(1, 3); Uniform = $true; RunOverrides = @{ 0 = $two }; FrenchRunParagraph = 0 },
        @{ Name = 'MergeSameParagraphCount'; Order = @(0, 1, 2, 3); FrenchRunParagraph = 1; RunOverrides = @{ 0 = $normal } },
        @{ Name = 'SplitSameParagraphCount'; Order = @(0, 1, 2, 3); Uniform = $true; MixedParagraphs = @(0); FrenchRunParagraph = 1 },
        @{ Name = 'SplitHiddenSameParagraphCount'; Order = @(0, 1, 2, 3); Uniform = $true; RunOverrides = @{ 1 = $two }; FrenchRunParagraph = 1 }
    )
    foreach ($profile in $profiles) {
        $overrides = @{}
        if ($profile.ContainsKey('Uniform')) {
            for ($index = 0; $index -lt $profile.Order.Count; $index++) { $overrides[$index] = $normal }
            if ($profile.ContainsKey('MixedParagraphs')) {
                foreach ($index in $profile.MixedParagraphs) { $overrides.Remove($index) }
            }
        }
        if ($profile.ContainsKey('RunOverrides')) {
            foreach ($index in $profile.RunOverrides.Keys) { $overrides[$index] = $profile.RunOverrides[$index] }
        }
        $profile.RunOverrides = $overrides
        $profile.MixedRuns = $true
        $profile.Text = ($profile.Order | ForEach-Object { $texts[$_] }) -join "`r"
        $profile
    }
}
