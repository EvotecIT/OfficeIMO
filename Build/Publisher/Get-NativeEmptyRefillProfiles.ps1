# Independently authored final states distinguish an empty paragraph's body typing from its CR role.
function Get-NativeEmptyRefillProfiles {
    $profiles = @(
        @{ Name = 'EmptyBaseline'; Order = @(0, 1, 2, 3); Text = "First native paragraph`r`rThird native paragraph`r" },
        @{ Name = 'FillInteriorDropFirst'; Order = @(1, 2, 3); Text = "Filled interior paragraph`rThird native paragraph`r" },
        @{ Name = 'FillBothDropMiddle'; Order = @(0, 1, 3); Text = "First native paragraph`rFilled interior paragraph`rFilled final paragraph" },
        @{ Name = 'FillFinalOnly'; Order = @(3); Text = 'Filled final paragraph' },
        @{ Name = 'FillInteriorAsFinal'; Order = @(1); Text = 'Filled interior paragraph' },
        @{ Name = 'FillAndReorder'; Order = @(3, 0, 1); Text = "Filled final paragraph`rFirst native paragraph`rFilled interior paragraph" },
        @{ Name = 'FillDuplicates'; Order = @(0, 1, 1, 2, 3); Text = "First native paragraph`rFirst fill`rSecond longer fill`rThird native paragraph`r" },
        @{ Name = 'AmbiguousEmptyLanguage'; Order = @(1, 1, 3); Text = "`r`r"; FrenchParagraph = 1 },
        @{ Name = 'EquivalentEmptyTemplates'; Order = @(1, 1, 3); Text = "`r`r" },
        @{ Name = 'FillEquivalentTemplate'; Order = @(1, 3); Text = "Shared template fill`r" },
        @{ Name = 'FillUnicodeInterior'; Order = @(1, 2, 3); Text = "Filled Unicode Ω paragraph`rThird native paragraph`r" }
    )
    $profiles
}
